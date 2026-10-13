using System;
using System.Collections.Generic;
using System.Linq;

namespace PstSplitter.Pst
{
    /// <summary>Builds the blocks of a heap-on-node ([MS-PST] 2.3.1).</summary>
    public sealed class HeapBuilder
    {
        public const int MaxAlloc = 3580;

        sealed class Block
        {
            public readonly List<byte[]> Allocs = new List<byte[]>();
            public int Used;
        }

        readonly List<Block> blocks = new List<Block> { new Block() };

        static int HeaderSize(int i) => i == 0 ? 12 : (i >= 8 && (i - 8) % 128 == 0) ? 66 : 2;

        static int BlockLength(int index, int used, int allocCount) =>
            ((HeaderSize(index) + used + 1) & ~1) + 4 + 2 * (allocCount + 1);

        public uint Add(byte[] data)
        {
            if (data.Length > MaxAlloc) throw new ArgumentException("Heap allocation too large.");
            int i = blocks.Count - 1;
            var b = blocks[i];
            if (b.Allocs.Count >= 2047 || BlockLength(i, b.Used + data.Length, b.Allocs.Count + 1) > PstFile.MaxBlockData)
            {
                blocks.Add(b = new Block());
                i++;
            }
            b.Allocs.Add(data);
            b.Used += data.Length;
            return (uint)((i << 16) | (b.Allocs.Count << 5));
        }

        public uint Reserve(int size) => Add(new byte[size]);

        public void Set(uint hid, byte[] data)
        {
            var b = blocks[(int)(hid >> 16)];
            int index = (int)((hid >> 5) & 0x7FF) - 1;
            if (b.Allocs[index].Length != data.Length) throw new InvalidOperationException("Heap allocation size changed.");
            b.Allocs[index] = data;
        }

        public List<byte[]> Build(byte clientSig, uint userRoot)
        {
            var result = new List<byte[]>();
            for (int i = 0; i < blocks.Count; i++)
            {
                var b = blocks[i];
                int len = BlockLength(i, b.Used, b.Allocs.Count);
                var data = new byte[len];
                int o = HeaderSize(i);
                var offsets = new List<int>();
                foreach (var a in b.Allocs)
                {
                    offsets.Add(o);
                    Buffer.BlockCopy(a, 0, data, o, a.Length);
                    o += a.Length;
                }
                offsets.Add(o);
                int ibHnpm = (o + 1) & ~1;
                BitConverter.GetBytes((ushort)ibHnpm).CopyTo(data, 0);
                if (i == 0)
                {
                    data[2] = 0xEC;
                    data[3] = clientSig;
                    BitConverter.GetBytes(userRoot).CopyTo(data, 4);
                }
                BitConverter.GetBytes((ushort)b.Allocs.Count).CopyTo(data, ibHnpm);
                BitConverter.GetBytes((ushort)0).CopyTo(data, ibHnpm + 2);
                for (int k = 0; k < offsets.Count; k++) BitConverter.GetBytes((ushort)offsets[k]).CopyTo(data, ibHnpm + 4 + k * 2);
                result.Add(data);
            }
            return result;
        }
    }

    public static class LtpWriter
    {
        /// <summary>
        /// Rebuilds a table context with only the given rows (raw row bytes in the source table's layout;
        /// fixed-size cells may have been changed). Variable-size cells are copied from the source table.
        /// Returns the new node's (bidData, bidSub).
        /// </summary>
        public static KeyValuePair<ulong, ulong> RebuildTable(PstWriter w, TableContext src, IList<byte[]> rows, Func<byte[], byte[]> transform = null)
        {
            var heap = new HeapBuilder();
            uint hidBth = heap.Reserve(8);
            int cCols = src.Columns.Count;
            uint hidInfo = heap.Reserve(22 + 8 * cCols);
            var subnodes = new List<SubnodeEntry>();
            uint nextSubIndex = 1;
            uint NewSubNid() => Nid.Make(0x1F, nextSubIndex++);

            var varCols = src.Columns.Where(c => { int f = PropType.FixedSize(c.Type); return f == 0 || f > 8; }).ToList();
            var outRows = new List<byte[]>();
            foreach (var original in rows)
            {
                var row = (byte[])original.Clone();
                foreach (var c in varCols)
                {
                    if (!src.CellExists(row, c)) continue;
                    uint hnid = BitConverter.ToUInt32(row, c.Ib);
                    byte[] value;
                    try { value = src.Heap.Value(hnid); }
                    catch (PstFormatException)
                    {
                        // Unreadable cell: drop it rather than copy a broken reference.
                        row[src.Rgib[2] + c.Bit / 8] &= (byte)~(0x80 >> (c.Bit % 8));
                        BitConverter.GetBytes(0u).CopyTo(row, c.Ib);
                        continue;
                    }
                    if (transform != null) value = transform(value);
                    uint newHnid;
                    if (value.Length == 0) newHnid = 0;
                    else if (value.Length <= HeapBuilder.MaxAlloc) newHnid = heap.Add(value);
                    else
                    {
                        newHnid = NewSubNid();
                        subnodes.Add(new SubnodeEntry { Nid = newHnid, BidData = w.NewData(value) });
                    }
                    BitConverter.GetBytes(newHnid).CopyTo(row, c.Ib);
                }
                outRows.Add(row);
            }

            // Row matrix in a subnode, whole rows per block.
            uint hnidRows = 0;
            int rowSize = src.RowSize;
            if (outRows.Count > 0)
            {
                int perBlock = PstFile.MaxBlockData / rowSize;
                var payloads = new List<byte[]>();
                for (int i = 0; i < outRows.Count; i += perBlock)
                {
                    int n = Math.Min(perBlock, outRows.Count - i);
                    var p = new byte[n * rowSize];
                    for (int k = 0; k < n; k++) Buffer.BlockCopy(outRows[i + k], 0, p, k * rowSize, rowSize);
                    payloads.Add(p);
                }
                hnidRows = NewSubNid();
                subnodes.Add(new SubnodeEntry { Nid = hnidRows, BidData = w.NewDataTree(payloads) });
            }

            // Row index: row ID -> row index.
            var index = outRows.Select((r, i) => new KeyValuePair<uint, uint>(TableContext.RowId(r), (uint)i))
                .OrderBy(kv => kv.Key).ToList();
            uint hidRoot = BuildBth(heap, index.Select(kv => Concat(BitConverter.GetBytes(kv.Key), BitConverter.GetBytes(kv.Value))).ToList(), 4, out byte levels);
            var bthHeader = new byte[8];
            bthHeader[0] = 0xB5;
            bthHeader[1] = 4;
            bthHeader[2] = 4;
            bthHeader[3] = levels;
            BitConverter.GetBytes(hidRoot).CopyTo(bthHeader, 4);
            heap.Set(hidBth, bthHeader);

            var info = new byte[22 + 8 * cCols];
            info[0] = 0x7C;
            info[1] = (byte)cCols;
            for (int i = 0; i < 4; i++) BitConverter.GetBytes(src.Rgib[i]).CopyTo(info, 2 + i * 2);
            BitConverter.GetBytes(hidBth).CopyTo(info, 10);
            BitConverter.GetBytes(hnidRows).CopyTo(info, 14);
            for (int i = 0; i < cCols; i++)
            {
                var c = src.Columns[i];
                int o = 22 + i * 8;
                BitConverter.GetBytes(c.Tag).CopyTo(info, o);
                BitConverter.GetBytes(c.Ib).CopyTo(info, o + 4);
                info[o + 6] = c.Cb;
                info[o + 7] = c.Bit;
            }
            heap.Set(hidInfo, info);

            ulong bidData = w.NewDataTree(heap.Build(0x7C, hidInfo));
            ulong bidSub = w.NewSubnodeTree(subnodes);
            return new KeyValuePair<ulong, ulong>(bidData, bidSub);
        }

        static byte[] Concat(byte[] a, byte[] b)
        {
            var r = new byte[a.Length + b.Length];
            Buffer.BlockCopy(a, 0, r, 0, a.Length);
            Buffer.BlockCopy(b, 0, r, a.Length, b.Length);
            return r;
        }

        /// <summary>Builds a BTH over sorted leaf records and returns the root HID (0 if empty).</summary>
        static uint BuildBth(HeapBuilder heap, List<byte[]> records, int cbKey, out byte levels)
        {
            levels = 0;
            if (records.Count == 0) return 0;
            var level = records;
            while (true)
            {
                int recSize = level[0].Length;
                int perAlloc = HeapBuilder.MaxAlloc / recSize;
                var parents = new List<byte[]>();
                for (int i = 0; i < level.Count; i += perAlloc)
                {
                    int n = Math.Min(perAlloc, level.Count - i);
                    var a = new byte[n * recSize];
                    for (int k = 0; k < n; k++) Buffer.BlockCopy(level[i + k], 0, a, k * recSize, recSize);
                    uint hid = heap.Add(a);
                    var p = new byte[cbKey + 4];
                    Buffer.BlockCopy(level[i], 0, p, 0, cbKey);
                    BitConverter.GetBytes(hid).CopyTo(p, cbKey);
                    parents.Add(p);
                }
                if (parents.Count == 1) return BitConverter.ToUInt32(parents[0], cbKey);
                level = parents;
                levels++;
            }
        }

        /// <summary>
        /// Changes small inline property values (4 bytes or less) in a property context.
        /// Returns the modified heap blocks, or null when nothing changed.
        /// </summary>
        public static List<byte[]> PatchProperties(HeapNode heap, IDictionary<ushort, uint> values)
        {
            var blocks = heap.Blocks.Select(b => (byte[])b.Clone()).ToList();
            bool changed = false;
            var hdr = heap.Alloc(heap.UserRoot);
            int levels = hdr[3];
            uint root = BitConverter.ToUInt32(hdr, 4);
            void Walk(uint hid, int level)
            {
                heap.Locate(hid, out int block, out int start, out int length);
                var b = blocks[block];
                if (level > 0)
                {
                    for (int o = start; o + 6 <= start + length; o += 6) Walk(BitConverter.ToUInt32(b, o + 2), level - 1);
                    return;
                }
                for (int o = start; o + 8 <= start + length; o += 8)
                {
                    ushort id = BitConverter.ToUInt16(b, o);
                    ushort type = BitConverter.ToUInt16(b, o + 2);
                    int size = PropType.FixedSize(type);
                    if (size == 0 || size > 4 || !values.TryGetValue(id, out uint v)) continue;
                    if (BitConverter.ToUInt32(b, o + 4) == v) continue;
                    BitConverter.GetBytes(v).CopyTo(b, o + 4);
                    changed = true;
                }
            }
            if (root != 0) Walk(root, levels);
            return changed ? blocks : null;
        }

        /// <summary>Replaces every occurrence of a byte pattern in the heap blocks. Returns null when not found.</summary>
        public static List<byte[]> ReplaceBytes(IList<byte[]> source, byte[] find, byte[] replace)
        {
            var blocks = source.Select(b => (byte[])b.Clone()).ToList();
            bool changed = false;
            foreach (var b in blocks)
            {
                for (int i = 0; i + find.Length <= b.Length; i++)
                {
                    int k = 0;
                    while (k < find.Length && b[i + k] == find[k]) k++;
                    if (k < find.Length) continue;
                    Buffer.BlockCopy(replace, 0, b, i, replace.Length);
                    changed = true;
                    i += find.Length - 1;
                }
            }
            return changed ? blocks : null;
        }
    }
}
