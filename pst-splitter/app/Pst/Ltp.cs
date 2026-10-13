using System;
using System.Collections.Generic;
using System.Text;

namespace PstSplitter.Pst
{
    /// <summary>
    /// A heap-on-node ([MS-PST] 2.3.1): the node's data blocks, decoded, with access to allocations by HID.
    /// </summary>
    public sealed class HeapNode
    {
        public readonly PstFile File;
        public readonly List<ulong> BlockIds;
        public readonly List<byte[]> Blocks = new List<byte[]>();
        public readonly Dictionary<uint, SubnodeEntry> Subnodes;
        public byte ClientSig => Blocks[0][3];
        public uint UserRoot => BitConverter.ToUInt32(Blocks[0], 4);

        public HeapNode(PstFile file, ulong bidData, ulong bidSub)
            : this(file, bidData, file.ReadSubnodes(bidSub)) { }

        public HeapNode(PstFile file, ulong bidData, Dictionary<uint, SubnodeEntry> subnodes)
        {
            File = file;
            Subnodes = subnodes;
            BlockIds = file.DataBlockIds(bidData);
            foreach (var bid in BlockIds) Blocks.Add(file.ReadBlock(bid));
            if (Blocks.Count == 0 || Blocks[0].Length < 12 || Blocks[0][2] != 0xEC)
                throw new PstFormatException("A table or property list in the PST is damaged.");
        }

        /// <summary>Block index, start offset and length of an allocation.</summary>
        public void Locate(uint hid, out int block, out int start, out int length)
        {
            if ((hid & 0x1F) != 0) throw new PstFormatException("Expected a heap ID, found a node ID.");
            block = (int)(hid >> 16);
            int index = (int)((hid >> 5) & 0x7FF);
            if (block >= Blocks.Count || index == 0) throw new PstFormatException("A heap reference in the PST is out of range.");
            var b = Blocks[block];
            int ibHnpm = BitConverter.ToUInt16(b, 0);
            int cAlloc = BitConverter.ToUInt16(b, ibHnpm);
            if (index > cAlloc) throw new PstFormatException("A heap reference in the PST is out of range.");
            start = BitConverter.ToUInt16(b, ibHnpm + 4 + (index - 1) * 2);
            int end = BitConverter.ToUInt16(b, ibHnpm + 4 + index * 2);
            length = end - start;
        }

        public byte[] Alloc(uint hid)
        {
            if (hid == 0) return new byte[0];
            Locate(hid, out int block, out int start, out int length);
            var r = new byte[length];
            Buffer.BlockCopy(Blocks[block], start, r, 0, length);
            return r;
        }

        /// <summary>Value bytes for an HNID: a heap allocation or a subnode's data.</summary>
        public byte[] Value(uint hnid)
        {
            if (hnid == 0) return new byte[0];
            if ((hnid & 0x1F) == 0) return Alloc(hnid);
            if (!Subnodes.TryGetValue(hnid, out var sn)) throw new PstFormatException("A value stored outside its table is missing from the PST.");
            return File.ReadData(sn.BidData);
        }

        /// <summary>Reads all leaf records of a BTH ([MS-PST] 2.3.2) as (key, data) pairs.</summary>
        public List<KeyValuePair<byte[], byte[]>> ReadBth(uint hidHeader)
        {
            var result = new List<KeyValuePair<byte[], byte[]>>();
            var hdr = Alloc(hidHeader);
            if (hdr.Length < 8 || hdr[0] != 0xB5) throw new PstFormatException("An index inside the PST is damaged.");
            int cbKey = hdr[1], cbEnt = hdr[2], levels = hdr[3];
            uint root = BitConverter.ToUInt32(hdr, 4);
            if (root != 0) ReadBthLevel(root, levels, cbKey, cbEnt, result);
            return result;
        }

        void ReadBthLevel(uint hid, int level, int cbKey, int cbEnt, List<KeyValuePair<byte[], byte[]>> result)
        {
            var data = Alloc(hid);
            if (level > 0)
            {
                int rec = cbKey + 4;
                for (int o = 0; o + rec <= data.Length; o += rec)
                    ReadBthLevel(BitConverter.ToUInt32(data, o + cbKey), level - 1, cbKey, cbEnt, result);
            }
            else
            {
                int rec = cbKey + cbEnt;
                for (int o = 0; o + rec <= data.Length; o += rec)
                {
                    var k = new byte[cbKey];
                    var v = new byte[cbEnt];
                    Buffer.BlockCopy(data, o, k, 0, cbKey);
                    Buffer.BlockCopy(data, o + cbKey, v, 0, cbEnt);
                    result.Add(new KeyValuePair<byte[], byte[]>(k, v));
                }
            }
        }
    }

    public static class PropType
    {
        public const ushort Int16 = 0x0002, Int32 = 0x0003, Float = 0x0004, Double = 0x0005, Currency = 0x0006,
            AppTime = 0x0007, Error = 0x000A, Boolean = 0x000B, Object = 0x000D, Int64 = 0x0014,
            String8 = 0x001E, String = 0x001F, Time = 0x0040, Guid = 0x0048, Binary = 0x0102;

        /// <summary>Size of a fixed-size type, or 0 for variable-size types.</summary>
        public static int FixedSize(ushort type)
        {
            switch (type)
            {
                case Int16: return 2;
                case Int32: case Float: case Error: return 4;
                case Boolean: return 1;
                case Double: case Currency: case AppTime: case Int64: case Time: return 8;
                case Guid: return 16;
                default: return 0;
            }
        }
    }

    /// <summary>A property context ([MS-PST] 2.3.3): the properties of a message, folder, attachment or store.</summary>
    public sealed class PropertyContext
    {
        public readonly HeapNode Heap;
        public readonly Dictionary<ushort, PropRecord> Props = new Dictionary<ushort, PropRecord>();

        public struct PropRecord
        {
            public ushort Id, Type;
            public uint ValueHnid;
        }

        public PropertyContext(HeapNode heap)
        {
            Heap = heap;
            if (heap.ClientSig != 0xBC) throw new PstFormatException("Expected a property list in the PST.");
            foreach (var kv in heap.ReadBth(heap.UserRoot))
            {
                var r = new PropRecord
                {
                    Id = BitConverter.ToUInt16(kv.Key, 0),
                    Type = BitConverter.ToUInt16(kv.Value, 0),
                    ValueHnid = BitConverter.ToUInt32(kv.Value, 2),
                };
                Props[r.Id] = r;
            }
        }

        public static PropertyContext Open(PstFile file, ulong bidData, ulong bidSub)
        {
            if (bidData == 0) return null;
            return new PropertyContext(new HeapNode(file, bidData, bidSub));
        }

        public bool Has(ushort id) => Props.ContainsKey(id);

        public byte[] Bytes(ushort id)
        {
            if (!Props.TryGetValue(id, out var r)) return null;
            int fixedSize = PropType.FixedSize(r.Type);
            if (fixedSize > 0 && fixedSize <= 4) return BitConverter.GetBytes(r.ValueHnid);
            return Heap.Value(r.ValueHnid);
        }

        public int? Int(ushort id)
        {
            if (!Props.TryGetValue(id, out var r)) return null;
            if (r.Type == PropType.Int32 || r.Type == PropType.Int16 || r.Type == PropType.Boolean) return (int)r.ValueHnid;
            return null;
        }

        public DateTime? Time(ushort id)
        {
            if (!Props.TryGetValue(id, out var r) || r.Type != PropType.Time) return null;
            var b = Heap.Value(r.ValueHnid);
            if (b.Length < 8) return null;
            long ft = BitConverter.ToInt64(b, 0);
            if (ft <= 0 || ft >= DateTime.MaxValue.ToFileTimeUtc()) return null;
            return DateTime.FromFileTimeUtc(ft).ToLocalTime();
        }

        public string String(ushort id)
        {
            if (!Props.TryGetValue(id, out var r)) return null;
            if (r.Type == PropType.String) return Encoding.Unicode.GetString(Heap.Value(r.ValueHnid)).TrimEnd('\0');
            if (r.Type == PropType.String8) return Encoding.GetEncoding(28591).GetString(Heap.Value(r.ValueHnid)).TrimEnd('\0');
            return null;
        }
    }

    /// <summary>A table context ([MS-PST] 2.3.4), e.g. a folder's contents or a message's recipients.</summary>
    public sealed class TableContext
    {
        public readonly HeapNode Heap;
        public readonly List<Column> Columns = new List<Column>();
        public readonly ushort[] Rgib = new ushort[4];
        public int RowSize => Rgib[3];
        /// <summary>Rows in row-index order, each the raw row bytes.</summary>
        public readonly List<byte[]> Rows = new List<byte[]>();

        public struct Column
        {
            public uint Tag;
            public ushort Ib;
            public byte Cb;
            public byte Bit;
            public ushort Type => (ushort)(Tag & 0xFFFF);
            public ushort Id => (ushort)(Tag >> 16);
        }

        public TableContext(HeapNode heap)
        {
            Heap = heap;
            if (heap.ClientSig != 0x7C) throw new PstFormatException("Expected a table in the PST.");
            var info = heap.Alloc(heap.UserRoot);
            if (info[0] != 0x7C) throw new PstFormatException("A table in the PST is damaged.");
            int cCols = info[1];
            for (int i = 0; i < 4; i++) Rgib[i] = BitConverter.ToUInt16(info, 2 + i * 2);
            uint hidRowIndex = BitConverter.ToUInt32(info, 10);
            uint hnidRows = BitConverter.ToUInt32(info, 14);
            for (int i = 0; i < cCols; i++)
            {
                int o = 22 + i * 8;
                Columns.Add(new Column { Tag = BitConverter.ToUInt32(info, o), Ib = BitConverter.ToUInt16(info, o + 4), Cb = info[o + 6], Bit = info[o + 7] });
            }

            var index = heap.ReadBth(hidRowIndex);
            if (index.Count == 0 || hnidRows == 0) return;

            var matrix = new List<byte[]>();
            int rowSize = RowSize;
            if ((hnidRows & 0x1F) == 0)
            {
                var all = heap.Alloc(hnidRows);
                for (int o = 0; o + rowSize <= all.Length; o += rowSize) matrix.Add(Slice(all, o, rowSize));
            }
            else
            {
                if (!heap.Subnodes.TryGetValue(hnidRows, out var sn)) throw new PstFormatException("A table's rows are missing from the PST.");
                foreach (var bid in heap.File.DataBlockIds(sn.BidData))
                {
                    var b = heap.File.ReadBlock(bid);
                    for (int o = 0; o + rowSize <= b.Length; o += rowSize) matrix.Add(Slice(b, o, rowSize));
                }
            }

            // Keep the rows the index refers to, in row-index order.
            var order = new SortedDictionary<int, byte[]>();
            foreach (var kv in index)
            {
                int rowIndex = (int)BitConverter.ToUInt32(kv.Value, 0);
                if (rowIndex < matrix.Count) order[rowIndex] = matrix[rowIndex];
            }
            Rows.AddRange(order.Values);
        }

        static byte[] Slice(byte[] b, int o, int n)
        {
            var r = new byte[n];
            Buffer.BlockCopy(b, o, r, 0, n);
            return r;
        }

        public static TableContext Open(PstFile file, ulong bidData, ulong bidSub)
        {
            if (bidData == 0) return null;
            return new TableContext(new HeapNode(file, bidData, bidSub));
        }

        public static TableContext Open(PstFile file, SubnodeEntry sn) => Open(file, sn.BidData, sn.BidSub);

        public static uint RowId(byte[] row) => BitConverter.ToUInt32(row, 0);

        public int ColumnIndex(uint tag)
        {
            for (int i = 0; i < Columns.Count; i++) if (Columns[i].Tag == tag) return i;
            return -1;
        }

        public bool CellExists(byte[] row, Column c) => (row[Rgib[2] + c.Bit / 8] & (0x80 >> (c.Bit % 8))) != 0;

        /// <summary>The value of a cell: inline bytes for small fixed types, otherwise the referenced data.</summary>
        public byte[] CellBytes(byte[] row, Column c)
        {
            if (!CellExists(row, c)) return null;
            int fixedSize = PropType.FixedSize(c.Type);
            if (fixedSize > 0 && fixedSize <= 8)
            {
                var r = new byte[c.Cb];
                Buffer.BlockCopy(row, c.Ib, r, 0, c.Cb);
                return r;
            }
            return Heap.Value(BitConverter.ToUInt32(row, c.Ib));
        }

        public string CellString(byte[] row, uint tag)
        {
            int i = ColumnIndex(tag);
            if (i < 0) return null;
            var b = CellBytes(row, Columns[i]);
            if (b == null) return null;
            return (tag & 0xFFFF) == PropType.String ? Encoding.Unicode.GetString(b).TrimEnd('\0') : Encoding.GetEncoding(28591).GetString(b).TrimEnd('\0');
        }
    }
}
