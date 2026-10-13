using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PstSplitter.Pst
{
    /// <summary>
    /// Writes a new Unicode PST made of nodes and blocks copied from a source PST plus newly built blocks.
    /// Copied blocks keep their IDs and contents; only their position in the file changes.
    /// Lays out the node and block B-trees, allocation maps and header ([MS-PST] 2.2).
    /// </summary>
    public sealed class PstWriter
    {
        const long AMapFirst = 0x4400;
        const long AMapData = 496L * 8 * 64;   // bytes covered by one AMap page
        const long DListOffset = 0x4200;
        const int FMapFirst = 128, FMapEvery = 496, FPMapFirst = 128 * 64, FPMapEvery = 496 * 64;

        sealed class OutBlock
        {
            public ulong Bid;
            public ushort Cb;
            public long SourceIb = -1;   // copied from the source file
            public byte[] Bytes;         // or newly built (already encoded)
            public long Ib;
            public int Refs;
        }

        readonly PstFile src;
        readonly Dictionary<ulong, OutBlock> blocks = new Dictionary<ulong, OutBlock>();
        readonly SortedDictionary<ulong, NodeEntry> nodes = new SortedDictionary<ulong, NodeEntry>();
        ulong nextBlockIndex;
        public long BytesWritten { get; private set; }

        public PstWriter(PstFile source)
        {
            src = source;
            nextBlockIndex = BitConverter.ToUInt64(src.Header, PstFile.OffBidNextB) >> 2;
        }

        // ---- building --------------------------------------------------------

        public void AddNode(ulong nid, ulong bidData, ulong bidSub, uint nidParent)
        {
            nodes[nid] = new NodeEntry { Nid = nid, BidData = bidData, BidSub = bidSub, NidParent = nidParent };
        }

        /// <summary>Includes a source block and everything below it (data trees, subnode trees).</summary>
        public ulong CopySource(ulong bid)
        {
            if (bid == 0) return 0;
            ulong key = bid & ~1UL;
            if (blocks.ContainsKey(key)) return bid;
            var e = src.GetBlock(bid);
            blocks[key] = new OutBlock { Bid = e.Bid, Cb = e.Cb, SourceIb = (long)e.Ib };
            if (PstFile.IsInternal(bid))
            {
                var raw = src.ReadRawBlock(bid);
                foreach (var child in InternalChildren(raw)) CopySource(child);
            }
            return bid;
        }

        static IEnumerable<ulong> InternalChildren(byte[] raw)
        {
            int cEnt = BitConverter.ToUInt16(raw, 2);
            if (raw[0] == 1)
            {
                for (int i = 0; i < cEnt; i++) yield return BitConverter.ToUInt64(raw, 8 + i * 8);
            }
            else if (raw[0] == 2 && raw[1] == 0)
            {
                for (int i = 0; i < cEnt; i++)
                {
                    yield return BitConverter.ToUInt64(raw, 8 + i * 24 + 8);
                    ulong sub = BitConverter.ToUInt64(raw, 8 + i * 24 + 16);
                    if (sub != 0) yield return sub;
                }
            }
            else if (raw[0] == 2)
            {
                for (int i = 0; i < cEnt; i++) yield return BitConverter.ToUInt64(raw, 8 + i * 16 + 8);
            }
        }

        ulong NewBid(bool internalBlock)
        {
            ulong bid = (nextBlockIndex << 2) | (internalBlock ? 2UL : 0UL);
            nextBlockIndex++;
            return bid;
        }

        ulong AddNewBlock(byte[] data, bool internalBlock)
        {
            if (data.Length == 0 || data.Length > PstFile.MaxBlockData) throw new InvalidOperationException("Block size out of range.");
            ulong bid = NewBid(internalBlock);
            if (!internalBlock) PstCrypto.Encode(data, data.Length, src.CryptMethod, bid);
            blocks[bid] = new OutBlock { Bid = bid, Cb = (ushort)data.Length, Bytes = data };
            return bid;
        }

        /// <summary>A data tree whose data blocks are exactly the given payloads (used for heaps).</summary>
        public ulong NewDataTree(IList<byte[]> payloads)
        {
            if (payloads.Count == 0) return 0;
            var bids = new List<ulong>();
            long total = 0;
            foreach (var p in payloads)
            {
                bids.Add(AddNewBlock((byte[])p.Clone(), false));
                total += p.Length;
            }
            if (bids.Count == 1) return bids[0];
            const int perX = (PstFile.MaxBlockData - 8) / 8;
            if (bids.Count <= perX) return XBlock(1, bids, total);
            var xs = new List<ulong>();
            long done = 0;
            for (int i = 0; i < bids.Count; i += perX)
            {
                var part = bids.Skip(i).Take(perX).ToList();
                long size = 0;
                for (int k = i; k < i + part.Count; k++) size += payloads[k].Length;
                xs.Add(XBlock(1, part, size));
                done += size;
            }
            if (xs.Count > perX) throw new InvalidOperationException("Value too large for a PST.");
            return XBlock(2, xs, total);
        }

        ulong XBlock(byte level, List<ulong> children, long total)
        {
            var b = new byte[8 + 8 * children.Count];
            b[0] = 1;
            b[1] = level;
            BitConverter.GetBytes((ushort)children.Count).CopyTo(b, 2);
            BitConverter.GetBytes((uint)total).CopyTo(b, 4);
            for (int i = 0; i < children.Count; i++) BitConverter.GetBytes(children[i]).CopyTo(b, 8 + i * 8);
            return AddNewBlock(b, true);
        }

        /// <summary>A data tree holding arbitrary bytes.</summary>
        public ulong NewData(byte[] data)
        {
            var parts = new List<byte[]>();
            for (int o = 0; o < data.Length; o += PstFile.MaxBlockData)
            {
                int n = Math.Min(PstFile.MaxBlockData, data.Length - o);
                var p = new byte[n];
                Buffer.BlockCopy(data, o, p, 0, n);
                parts.Add(p);
            }
            return NewDataTree(parts);
        }

        /// <summary>A subnode tree (SLBLOCK, plus SIBLOCK when large).</summary>
        public ulong NewSubnodeTree(IEnumerable<SubnodeEntry> entries)
        {
            var list = entries.OrderBy(e => e.Nid).ToList();
            if (list.Count == 0) return 0;
            const int perSl = (PstFile.MaxBlockData - 8) / 24;
            const int perSi = (PstFile.MaxBlockData - 8) / 16;
            var leaves = new List<KeyValuePair<uint, ulong>>();
            for (int i = 0; i < list.Count; i += perSl)
            {
                var part = list.Skip(i).Take(perSl).ToList();
                var b = new byte[8 + 24 * part.Count];
                b[0] = 2;
                BitConverter.GetBytes((ushort)part.Count).CopyTo(b, 2);
                for (int k = 0; k < part.Count; k++)
                {
                    BitConverter.GetBytes((ulong)part[k].Nid).CopyTo(b, 8 + k * 24);
                    BitConverter.GetBytes(part[k].BidData).CopyTo(b, 8 + k * 24 + 8);
                    BitConverter.GetBytes(part[k].BidSub).CopyTo(b, 8 + k * 24 + 16);
                }
                leaves.Add(new KeyValuePair<uint, ulong>(part[0].Nid, AddNewBlock(b, true)));
            }
            if (leaves.Count == 1) return leaves[0].Value;
            if (leaves.Count > perSi) throw new InvalidOperationException("Too many attachments or values in one item.");
            var si = new byte[8 + 16 * leaves.Count];
            si[0] = 2;
            si[1] = 1;
            BitConverter.GetBytes((ushort)leaves.Count).CopyTo(si, 2);
            for (int k = 0; k < leaves.Count; k++)
            {
                BitConverter.GetBytes((ulong)leaves[k].Key).CopyTo(si, 8 + k * 16);
                BitConverter.GetBytes(leaves[k].Value).CopyTo(si, 8 + k * 16 + 8);
            }
            return AddNewBlock(si, true);
        }

        // ---- writing ---------------------------------------------------------

        static long IntervalStart(long i) => AMapFirst + i * AMapData;
        static bool HasPMap(long i) => i % 8 == 0;
        static bool HasFMap(long i) => i >= FMapFirst && (i - FMapFirst) % FMapEvery == 0;
        static bool HasFPMap(long i) => i >= FPMapFirst && (i - FPMapFirst) % FPMapEvery == 0;
        static int ReservedPages(long i) => HasFPMap(i) ? 4 : HasFMap(i) ? 3 : HasPMap(i) ? 2 : 1;

        long cursor;

        long Allocate(long size, long align)
        {
            cursor = (cursor + align - 1) / align * align;
            long i = (cursor - AMapFirst) / AMapData;
            if (cursor < IntervalStart(i) + ReservedPages(i) * 512L) cursor = IntervalStart(i) + ReservedPages(i) * 512L;
            if (cursor + size > IntervalStart(i + 1))
            {
                i++;
                cursor = IntervalStart(i) + ReservedPages(i) * 512L;
            }
            long pos = cursor;
            cursor += size;
            return pos;
        }

        static int BlockSpan(int cb) => (cb + 16 + 63) / 64 * 64;

        public void Write(string path, Func<bool> shouldCancel = null, Action<long, long> progress = null)
        {
            // Reference counts: Outlook stores (number of references + 1).
            foreach (var b in blocks.Values) b.Refs = 1;
            void Ref(ulong bid)
            {
                if (bid == 0) return;
                if (!blocks.TryGetValue(bid & ~1UL, out var b)) throw new InvalidOperationException(string.Format("Block {0:X} is referenced but not included.", bid));
                b.Refs++;
            }
            foreach (var n in nodes.Values) { Ref(n.BidData); Ref(n.BidSub); }
            foreach (var b in blocks.Values.ToList())
            {
                if (!PstFile.IsInternal(b.Bid)) continue;
                var raw = b.Bytes ?? src.ReadRawBlock(b.Bid);
                foreach (var c in InternalChildren(raw)) Ref(c);
            }

            // Lay out blocks (copied ones in source order so the source is read sequentially), then pages.
            cursor = IntervalStart(0) + ReservedPages(0) * 512L;
            var ordered = blocks.Values.Where(b => b.SourceIb >= 0).OrderBy(b => b.SourceIb)
                .Concat(blocks.Values.Where(b => b.SourceIb < 0)).ToList();
            foreach (var b in ordered) b.Ib = Allocate(BlockSpan(b.Cb), 64);

            ulong nextPage = BitConverter.ToUInt64(src.Header, PstFile.OffBidNextP);
            var pages = new List<KeyValuePair<long, byte[]>>();
            var bbtLeaves = blocks.Values.OrderBy(b => b.Bid).Select(b =>
            {
                var e = new byte[24];
                BitConverter.GetBytes(b.Bid).CopyTo(e, 0);
                BitConverter.GetBytes((ulong)b.Ib).CopyTo(e, 8);
                BitConverter.GetBytes(b.Cb).CopyTo(e, 16);
                BitConverter.GetBytes((ushort)Math.Min(b.Refs, ushort.MaxValue)).CopyTo(e, 18);
                return new KeyValuePair<ulong, byte[]>(b.Bid, e);
            }).ToList();
            var nbtLeaves = nodes.Values.Select(n =>
            {
                var e = new byte[32];
                BitConverter.GetBytes(n.Nid).CopyTo(e, 0);
                BitConverter.GetBytes(n.BidData).CopyTo(e, 8);
                BitConverter.GetBytes(n.BidSub).CopyTo(e, 16);
                BitConverter.GetBytes(n.NidParent).CopyTo(e, 24);
                return new KeyValuePair<ulong, byte[]>(n.Nid, e);
            }).ToList();
            var nbtRoot = BuildBTree(nbtLeaves, 0x81, 32, ref nextPage, pages);
            var bbtRoot = BuildBTree(bbtLeaves, 0x80, 24, ref nextPage, pages);
            long eof = cursor;

            // Allocation maps.
            long intervals = (eof - AMapFirst + AMapData - 1) / AMapData;
            var amaps = new byte[intervals][];
            for (long i = 0; i < intervals; i++)
            {
                amaps[i] = new byte[496];
                Mark(amaps, IntervalStart(i), 512);
                if (HasPMap(i)) Mark(amaps, IntervalStart(i) + 512, 512);
                // An FPMap interval reserves the page before the FPMap too, as Microsoft's outlook-pst does.
                if (HasFMap(i) || HasFPMap(i)) Mark(amaps, IntervalStart(i) + 1024, 512);
                if (HasFPMap(i)) Mark(amaps, IntervalStart(i) + 1536, 512);
            }
            foreach (var b in blocks.Values) Mark(amaps, b.Ib, BlockSpan(b.Cb));
            foreach (var p in pages) Mark(amaps, p.Key, 512);
            long freeBytes = 0;
            foreach (var m in amaps)
                foreach (var x in m)
                    for (int bit = 0; bit < 8; bit++) if ((x & (0x80 >> bit)) == 0) freeBytes += 64;

            using (var fs = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None, 1 << 20))
            {
                fs.SetLength(eof);

                // Blocks
                var buf = new byte[8192];
                long totalSpan = ordered.Sum(b => (long)BlockSpan(b.Cb)), doneSpan = 0;
                foreach (var b in ordered)
                {
                    if (shouldCancel != null && shouldCancel()) throw new OperationCanceledException();
                    doneSpan += BlockSpan(b.Cb);
                    progress?.Invoke(doneSpan, totalSpan);
                    byte[] data = b.Bytes ?? src.ReadAt(b.SourceIb, b.Cb);
                    int span = BlockSpan(b.Cb);
                    Array.Clear(buf, 0, span);
                    Buffer.BlockCopy(data, 0, buf, 0, b.Cb);
                    int t = span - 16;
                    BitConverter.GetBytes(b.Cb).CopyTo(buf, t);
                    BitConverter.GetBytes(PstCrypto.Signature((ulong)b.Ib, b.Bid)).CopyTo(buf, t + 2);
                    BitConverter.GetBytes(PstCrypto.Crc(data, 0, b.Cb)).CopyTo(buf, t + 4);
                    BitConverter.GetBytes(b.Bid).CopyTo(buf, t + 8);
                    fs.Position = b.Ib;
                    fs.Write(buf, 0, span);
                }

                foreach (var p in pages) { fs.Position = p.Key; fs.Write(p.Value, 0, 512); }

                for (long i = 0; i < intervals; i++)
                {
                    WriteMapPage(fs, IntervalStart(i), 0x84, amaps[i]);
                    if (HasPMap(i)) WriteMapPage(fs, IntervalStart(i) + 512, 0x83, Filled(0xFF));
                    if (HasFMap(i))
                    {
                        var fm = new byte[496];
                        for (int k = 0; k < 496 && i + k < intervals; k++) fm[k] = LongestFreeRun(amaps[i + k]);
                        WriteMapPage(fs, IntervalStart(i) + 1024, 0x82, fm);
                    }
                    if (HasFPMap(i)) WriteMapPage(fs, IntervalStart(i) + 1536, 0x85, Filled(0xFF));
                }

                // Density list (empty)
                var dl = new byte[512];
                dl[496] = dl[497] = 0x86;
                BitConverter.GetBytes(PstCrypto.Signature(DListOffset, nextPage)).CopyTo(dl, 498);
                BitConverter.GetBytes(PstCrypto.Crc(dl, 0, 496)).CopyTo(dl, 500);
                BitConverter.GetBytes(nextPage).CopyTo(dl, 504);
                fs.Position = DListOffset;
                fs.Write(dl, 0, 512);

                // Header
                var h = (byte[])src.Header.Clone();
                BitConverter.GetBytes(nextPage).CopyTo(h, PstFile.OffBidNextP);
                BitConverter.GetBytes(BitConverter.ToUInt32(h, PstFile.OffUnique) + 1).CopyTo(h, PstFile.OffUnique);
                int r = PstFile.OffRoot;
                BitConverter.GetBytes((ulong)eof).CopyTo(h, r + 4);
                BitConverter.GetBytes((ulong)IntervalStart(intervals - 1)).CopyTo(h, r + 12);
                BitConverter.GetBytes((ulong)freeBytes).CopyTo(h, r + 20);
                BitConverter.GetBytes(0UL).CopyTo(h, r + 28);
                BitConverter.GetBytes(nbtRoot.Key).CopyTo(h, r + 36);
                BitConverter.GetBytes((ulong)nbtRoot.Value).CopyTo(h, r + 44);
                BitConverter.GetBytes(bbtRoot.Key).CopyTo(h, r + 52);
                BitConverter.GetBytes((ulong)bbtRoot.Value).CopyTo(h, r + 60);
                h[r + 68] = 2; // VALID_AMAP2
                h[r + 69] = 0;
                h[r + 70] = h[r + 71] = 0;
                BitConverter.GetBytes(nextBlockIndex << 2).CopyTo(h, PstFile.OffBidNextB);
                BitConverter.GetBytes(PstCrypto.Crc(h, 8, 471)).CopyTo(h, 4);
                BitConverter.GetBytes(PstCrypto.Crc(h, 8, 516)).CopyTo(h, PstFile.OffCrcFull);
                fs.Position = 0;
                fs.Write(h, 0, h.Length);
                fs.Flush(true);
            }
            BytesWritten = eof;
        }

        static byte[] Filled(byte v)
        {
            var b = new byte[496];
            for (int i = 0; i < b.Length; i++) b[i] = v;
            return b;
        }

        static byte LongestFreeRun(byte[] map)
        {
            int best = 0, run = 0;
            foreach (var x in map)
                for (int bit = 0; bit < 8; bit++)
                {
                    if ((x & (0x80 >> bit)) == 0) { run++; if (run > best) best = run; }
                    else run = 0;
                }
            return (byte)Math.Min(best, 255);
        }

        static void WriteMapPage(Stream fs, long ib, byte type, byte[] bits)
        {
            var p = new byte[512];
            Buffer.BlockCopy(bits, 0, p, 0, 496);
            p[496] = p[497] = type;
            BitConverter.GetBytes(PstCrypto.Crc(p, 0, 496)).CopyTo(p, 500);
            BitConverter.GetBytes((ulong)ib).CopyTo(p, 504);
            fs.Position = ib;
            fs.Write(p, 0, 512);
        }

        static void Mark(byte[][] amaps, long ib, long size)
        {
            long x = ib - AMapFirst;
            var map = amaps[x / AMapData];
            long first = (x % AMapData) / 64;
            for (long bit = first; bit < first + size / 64; bit++) map[bit / 8] |= (byte)(0x80 >> (int)(bit % 8));
        }

        /// <summary>Builds a B-tree bottom-up and returns (root page bid, root page offset).</summary>
        KeyValuePair<ulong, long> BuildBTree(List<KeyValuePair<ulong, byte[]>> leafEntries, byte type, int leafSize, ref ulong nextPage, List<KeyValuePair<long, byte[]>> pages)
        {
            int level = 0;
            int entrySize = leafSize;
            var entries = leafEntries;
            while (true)
            {
                int perPage = 488 / entrySize;
                var parents = new List<KeyValuePair<ulong, byte[]>>();
                int count = Math.Max(1, (entries.Count + perPage - 1) / perPage);
                // Spread entries evenly across the pages of this level.
                for (int p = 0; p < count; p++)
                {
                    int from = (int)((long)entries.Count * p / count);
                    int to = (int)((long)entries.Count * (p + 1) / count);
                    var page = new byte[512];
                    for (int i = from; i < to; i++) Buffer.BlockCopy(entries[i].Value, 0, page, (i - from) * entrySize, entrySize);
                    page[488] = (byte)(to - from);
                    page[489] = (byte)perPage;
                    page[490] = (byte)entrySize;
                    page[491] = (byte)level;
                    long ib = Allocate(512, 512);
                    ulong pageBid = nextPage++;
                    page[496] = page[497] = type;
                    BitConverter.GetBytes(PstCrypto.Signature((ulong)ib, pageBid)).CopyTo(page, 498);
                    BitConverter.GetBytes(PstCrypto.Crc(page, 0, 496)).CopyTo(page, 500);
                    BitConverter.GetBytes(pageBid).CopyTo(page, 504);
                    pages.Add(new KeyValuePair<long, byte[]>(ib, page));

                    var e = new byte[24];
                    ulong firstKey = entries.Count > 0 ? entries[from].Key : 0;
                    BitConverter.GetBytes(firstKey).CopyTo(e, 0);
                    BitConverter.GetBytes(pageBid).CopyTo(e, 8);
                    BitConverter.GetBytes((ulong)ib).CopyTo(e, 16);
                    parents.Add(new KeyValuePair<ulong, byte[]>(firstKey, e));
                }
                if (parents.Count == 1)
                {
                    var root = parents[0].Value;
                    return new KeyValuePair<ulong, long>(BitConverter.ToUInt64(root, 8), (long)BitConverter.ToUInt64(root, 16));
                }
                entries = parents;
                entrySize = 24;
                level++;
            }
        }
    }
}
