using System;
using System.Collections.Generic;
using System.IO;

namespace PstSplitter.Pst
{
    public class PstFormatException : Exception
    {
        public PstFormatException(string message) : base(message) { }
    }

    public struct NodeEntry
    {
        public ulong Nid;
        public ulong BidData;
        public ulong BidSub;
        public uint NidParent;
    }

    public struct BlockEntry
    {
        public ulong Bid;
        public ulong Ib;
        public ushort Cb;
        public ushort RefCount;
    }

    /// <summary>
    /// Reads the node database (NDB) layer of a Unicode PST file: header, node and block B-trees,
    /// blocks, data trees and subnode trees. See [MS-PST] section 2.2.
    /// </summary>
    public sealed class PstFile : IDisposable
    {
        public const int PageSize = 512;
        public const int MaxBlockData = 8192 - 16;
        public const int HeaderSize = 564;

        // Unicode HEADER field offsets ([MS-PST] 2.2.2.6)
        public const int OffBidNextP = 32, OffUnique = 40, OffRgnid = 44, OffRoot = 180,
            OffRgbFM = 256, OffRgbFP = 384, OffCrypt = 513, OffBidNextB = 516, OffCrcFull = 524;

        readonly Stream stream;
        readonly object streamLock = new object();
        public byte[] Header { get; private set; }
        public byte CryptMethod { get; private set; }
        public Dictionary<ulong, NodeEntry> Nodes { get; } = new Dictionary<ulong, NodeEntry>();
        public Dictionary<ulong, BlockEntry> Blocks { get; } = new Dictionary<ulong, BlockEntry>();
        public long FileLength => stream.Length;

        readonly Action<int> onPagesRead;
        int pagesRead;

        /// <param name="onPagesRead">Called every 100 index pages while the B-trees load.</param>
        public PstFile(string path, Action<int> onPagesRead = null)
        {
            this.onPagesRead = onPagesRead;
            try
            {
                // Don't let anything (e.g. Outlook) write to the file while we read it.
                stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16);
            }
            catch (IOException) when (File.Exists(path))
            {
                throw new PstFormatException("The PST file is in use, probably by Outlook. Close Outlook (or remove this file from Outlook) and try again.");
            }
            try { Load(); }
            catch { stream.Dispose(); throw; }
        }

        public void Dispose() => stream.Dispose();

        void Load()
        {
            Header = ReadAt(0, HeaderSize);
            if (Header[0] != 0x21 || Header[1] != 0x42 || Header[2] != 0x44 || Header[3] != 0x4E)
                throw new PstFormatException("This is not an Outlook PST file.");
            if (Header[8] != 0x53 || Header[9] != 0x4D)
                throw new PstFormatException(Header[8] == 0x53 && Header[9] == 0x4F
                    ? "This is an OST (offline) file. Only PST files can be split."
                    : "This file is not an Outlook PST file.");
            ushort ver = BitConverter.ToUInt16(Header, 10);
            if (ver < 23)
                throw new PstFormatException("This PST uses the old ANSI format (Outlook 97-2002). Open it in a newer Outlook and export it to a new PST first.");
            if (ver >= 36)
                throw new PstFormatException("This file uses the 4K-page format, which is only used for OST files.");
            uint crc = PstCrypto.Crc(Header, 8, 516);
            if (crc != BitConverter.ToUInt32(Header, OffCrcFull))
                throw new PstFormatException("The PST header is damaged. Run the Inbox Repair Tool (scanpst.exe) on it first.");
            CryptMethod = Header[OffCrypt];
            if (CryptMethod > 2) throw new PstFormatException("Unknown PST encryption type " + CryptMethod + ".");

            ulong nbtIb = BitConverter.ToUInt64(Header, OffRoot + 44);
            ulong bbtIb = BitConverter.ToUInt64(Header, OffRoot + 60);
            ReadBTree(nbtIb, 0x81, true, -1);
            ReadBTree(bbtIb, 0x80, false, -1);
        }

        public byte[] ReadAt(long offset, int count)
        {
            var buf = new byte[count];
            lock (streamLock)
            {
                stream.Position = offset;
                int read = 0;
                while (read < count)
                {
                    int n = stream.Read(buf, read, count - read);
                    if (n <= 0) throw new PstFormatException("The PST file is shorter than expected. It may be damaged or still being copied.");
                    read += n;
                }
            }
            return buf;
        }

        void ReadBTree(ulong ib, byte pageType, bool nodes, int expectedLevel)
        {
            var page = ReadAt((long)ib, PageSize);
            if (++pagesRead % 100 == 0) onPagesRead?.Invoke(pagesRead);
            if (page[496] != pageType || page[497] != pageType)
                throw new PstFormatException("A B-tree page in the PST is damaged. Run the Inbox Repair Tool (scanpst.exe) on it first.");
            int cEnt = page[488], cbEnt = page[490], cLevel = page[491];
            if (expectedLevel >= 0 && cLevel != expectedLevel)
                throw new PstFormatException("The PST's index is damaged. Run the Inbox Repair Tool (scanpst.exe) on it first.");
            for (int i = 0; i < cEnt; i++)
            {
                int o = i * cbEnt;
                if (cLevel > 0)
                {
                    ulong childIb = BitConverter.ToUInt64(page, o + 16);
                    ReadBTree(childIb, pageType, nodes, cLevel - 1);
                }
                else if (nodes)
                {
                    var e = new NodeEntry
                    {
                        Nid = BitConverter.ToUInt64(page, o),
                        BidData = BitConverter.ToUInt64(page, o + 8),
                        BidSub = BitConverter.ToUInt64(page, o + 16),
                        NidParent = BitConverter.ToUInt32(page, o + 24),
                    };
                    Nodes[e.Nid] = e;
                }
                else
                {
                    var e = new BlockEntry
                    {
                        Bid = BitConverter.ToUInt64(page, o),
                        Ib = BitConverter.ToUInt64(page, o + 8),
                        Cb = BitConverter.ToUInt16(page, o + 16),
                        RefCount = BitConverter.ToUInt16(page, o + 18),
                    };
                    Blocks[e.Bid & ~1UL] = e;
                }
            }
        }

        public static bool IsInternal(ulong bid) => (bid & 2) != 0;

        public BlockEntry GetBlock(ulong bid)
        {
            if (!Blocks.TryGetValue(bid & ~1UL, out var e))
                throw new PstFormatException(string.Format("Block {0:X} is missing from the PST. Run the Inbox Repair Tool (scanpst.exe) on it first.", bid));
            return e;
        }

        /// <summary>Raw block bytes as stored (still encoded for data blocks), without the trailer.</summary>
        public byte[] ReadRawBlock(ulong bid)
        {
            var e = GetBlock(bid);
            return ReadAt((long)e.Ib, e.Cb);
        }

        /// <summary>Block contents, decoded if it's an external data block.</summary>
        public byte[] ReadBlock(ulong bid)
        {
            var data = ReadRawBlock(bid);
            if (!IsInternal(bid)) PstCrypto.Decode(data, data.Length, CryptMethod, bid);
            return data;
        }

        /// <summary>The data blocks making up a node's data (expands XBLOCK/XXBLOCK trees).</summary>
        public List<ulong> DataBlockIds(ulong bid)
        {
            var list = new List<ulong>();
            if (bid == 0) return list;
            CollectDataBlocks(bid, list);
            return list;
        }

        void CollectDataBlocks(ulong bid, List<ulong> list)
        {
            if (!IsInternal(bid)) { list.Add(bid); return; }
            var b = ReadRawBlock(bid);
            if (b[0] != 1) throw new PstFormatException("Unexpected block type in a data tree.");
            int cEnt = BitConverter.ToUInt16(b, 2);
            for (int i = 0; i < cEnt; i++)
                CollectDataBlocks(BitConverter.ToUInt64(b, 8 + i * 8), list);
        }

        /// <summary>All block IDs a data tree uses, including the XBLOCK/XXBLOCK blocks themselves.</summary>
        public void CollectTreeBlocks(ulong bid, ICollection<ulong> into)
        {
            if (bid == 0) return;
            into.Add(bid);
            if (!IsInternal(bid)) return;
            var b = ReadRawBlock(bid);
            int cEnt = BitConverter.ToUInt16(b, 2);
            if (b[0] == 1)
            {
                for (int i = 0; i < cEnt; i++) CollectTreeBlocks(BitConverter.ToUInt64(b, 8 + i * 8), into);
            }
            else if (b[0] == 2)
            {
                if (b[1] == 0)
                    for (int i = 0; i < cEnt; i++)
                    {
                        CollectTreeBlocks(BitConverter.ToUInt64(b, 8 + i * 24 + 8), into);
                        CollectTreeBlocks(BitConverter.ToUInt64(b, 8 + i * 24 + 16), into);
                    }
                else
                    for (int i = 0; i < cEnt; i++) CollectTreeBlocks(BitConverter.ToUInt64(b, 8 + i * 16 + 8), into);
            }
        }

        public byte[] ReadData(ulong bid)
        {
            var ms = new MemoryStream();
            foreach (var id in DataBlockIds(bid))
            {
                var b = ReadBlock(id);
                ms.Write(b, 0, b.Length);
            }
            return ms.ToArray();
        }

        /// <summary>Subnode tree (SLBLOCK/SIBLOCK) as nid -> (bidData, bidSub).</summary>
        public Dictionary<uint, SubnodeEntry> ReadSubnodes(ulong bidSub)
        {
            var d = new Dictionary<uint, SubnodeEntry>();
            if (bidSub != 0) CollectSubnodes(bidSub, d);
            return d;
        }

        void CollectSubnodes(ulong bid, Dictionary<uint, SubnodeEntry> d)
        {
            var b = ReadRawBlock(bid);
            if (b[0] != 2) throw new PstFormatException("Unexpected block type in a subnode tree.");
            int cEnt = BitConverter.ToUInt16(b, 2);
            if (b[1] == 0)
            {
                for (int i = 0; i < cEnt; i++)
                {
                    int o = 8 + i * 24;
                    uint nid = (uint)BitConverter.ToUInt64(b, o);
                    d[nid] = new SubnodeEntry { Nid = nid, BidData = BitConverter.ToUInt64(b, o + 8), BidSub = BitConverter.ToUInt64(b, o + 16) };
                }
            }
            else
            {
                for (int i = 0; i < cEnt; i++) CollectSubnodes(BitConverter.ToUInt64(b, 8 + i * 16 + 8), d);
            }
        }
    }

    public struct SubnodeEntry
    {
        public uint Nid;
        public ulong BidData;
        public ulong BidSub;
    }

    public static class Nid
    {
        public const uint MessageStore = 0x21, NameToIdMap = 0x61, RootFolder = 0x122;
        public const uint TypeNormalFolder = 0x02, TypeSearchFolder = 0x03, TypeNormalMessage = 0x04,
            TypeAttachment = 0x05, TypeAssocMessage = 0x08, TypeHierarchyTable = 0x0D, TypeContentsTable = 0x0E,
            TypeAssocContentsTable = 0x0F, TypeSearchContentsTable = 0x10, TypeSearchUpdateQueue = 0x11;
        public const uint RecipientTable = 0x692, AttachmentTable = 0x671;

        public static uint Type(ulong nid) => (uint)nid & 0x1F;
        public static uint Index(ulong nid) => (uint)nid >> 5;
        public static uint Make(uint type, uint index) => (index << 5) | type;
    }
}
