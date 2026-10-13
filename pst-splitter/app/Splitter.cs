using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PstSplitter.Pst;

namespace PstSplitter
{
    /// <summary>
    /// Splits a PST by reading it directly and writing new PST files; Outlook is not needed.
    /// Each output is a copy of the source's structure that keeps only the chosen messages, the
    /// folders they're in and Outlook's standard folders. Messages are copied byte for byte.
    /// </summary>
    public class Splitter
    {
        public Action<string> Log = _ => { };
        public Action<int, int, string> Progress = (done, total, label) => { };
        public Func<bool> ShouldCancel = () => false;

        // Property IDs
        const ushort PidDisplayName = 0x3001, PidContentCount = 0x3602, PidContentUnread = 0x3603, PidSubfolders = 0x360A,
            PidRecordKey = 0x0FF9, PidIpmSubtree = 0x35E0, PidWastebasket = 0x35E3, PidFinder = 0x35E7,
            PidDeliveryTime = 0x0E06, PidSubmitTime = 0x0039, PidCreationTime = 0x3007, PidMessageFlags = 0x0E07,
            PidSenderSmtp = 0x5D01, PidSenderEmail = 0x0C1F, PidSenderAddrType = 0x0C1E, PidSenderName = 0x0C1A,
            PidSentRepSmtp = 0x5D02, PidSentRepEmail = 0x0065;
        const uint TagContentCount = 0x36020003, TagContentUnread = 0x36030003, TagSubfolders = 0x360A000B,
            TagMessageFlags = 0x0E070003, TagSmtp = 0x39FE001F, TagEmail = 0x3003001F;

        class FolderInfo
        {
            public ulong Nid;
            public ulong Parent;
            public string Name;
            public bool IsSearch;
            public string[] Path;          // below the IPM subtree, or null if outside it
            public bool Special;           // never removed from an output
            public readonly List<FolderInfo> Children = new List<FolderInfo>();
        }

        class MessageInfo
        {
            public ulong Nid;
            public FolderInfo Folder;
            public long Bytes;
            public bool Unread;
        }

        class Group
        {
            public string Key;
            public int Part;
            public List<MessageInfo> Messages = new List<MessageInfo>();
            public long Bytes;
            public SplitFile File;
        }

        PstFile pst;
        SplitOptions opts;
        Dictionary<ulong, FolderInfo> folders;
        ushort? appointmentStartProp;
        long skeletonBytes;

        public SplitResult Run(SplitOptions options, bool preview)
        {
            opts = options;
            var source = Path.GetFullPath(opts.Source);
            var outDir = Path.GetFullPath(opts.Output);
            var result = new SplitResult { Preview = preview, Output = outDir };

            Log("Reading " + source);
            using (pst = new PstFile(source))
            {
                LoadFolders();
                appointmentStartProp = FindNamedProperty(new Guid("00062002-0000-0000-C000-000000000046"), 0x820D);

                // Messages, in folder order
                var messages = new List<MessageInfo>();
                var tableNids = new HashSet<ulong>();
                foreach (var f in folders.Values.Where(f => !f.IsSearch && f.Path != null).OrderBy(f => string.Join("\\", f.Path), StringComparer.OrdinalIgnoreCase))
                {
                    ulong ctNid = Nid.Make(Nid.TypeContentsTable, Nid.Index(f.Nid));
                    var ct = OpenTable(ctNid);
                    if (ct == null) continue;
                    tableNids.Add(ctNid);
                    // Each item also costs its share of the folder's contents table.
                    long rowCost = ct.Rows.Count == 0 ? 0 : (TreeBytes(pst.Nodes[ctNid].BidData) + TreeBytes(pst.Nodes[ctNid].BidSub)) / ct.Rows.Count + 1;
                    int flagsCol = ct.ColumnIndex(TagMessageFlags);
                    foreach (var row in ct.Rows)
                    {
                        ulong nid = TableContext.RowId(row);
                        if (Nid.Type(nid) != Nid.TypeNormalMessage || !pst.Nodes.TryGetValue(nid, out var node)) continue;
                        bool unread = false;
                        if (flagsCol >= 0)
                        {
                            var flags = ct.CellBytes(row, ct.Columns[flagsCol]);
                            unread = flags != null && (BitConverter.ToInt32(flags, 0) & 1) == 0;
                        }
                        long bytes;
                        try { bytes = TreeBytes(node.BidData) + TreeBytes(node.BidSub) + rowCost; }
                        catch (PstFormatException ex)
                        {
                            result.Failed++;
                            Log(string.Format("  Skipping a damaged item in {0}: {1}", string.Join("\\", f.Path), ex.Message));
                            continue;
                        }
                        messages.Add(new MessageInfo { Nid = nid, Folder = f, Bytes = bytes, Unread = unread });
                    }
                }
                skeletonBytes = 0x4800 + pst.Nodes.Values.Where(n => Nid.Type(n.Nid) != Nid.TypeNormalMessage && !tableNids.Contains(n.Nid))
                    .Sum(n => TreeBytes(n.BidData) + TreeBytes(n.BidSub) + 64);
                SetFolderKeys(messages);
                Log(string.Format("Found {0} folders and {1:N0} items", folders.Values.Count(f => f.Path != null && f.Path.Length > 0), messages.Count));

                // Group
                var groups = new Dictionary<string, List<Group>>();
                var order = new List<Group>();
                long maxBytes = (long)opts.MaxSizeMB * 1024 * 1024;
                string baseName = SafeName(Path.GetFileNameWithoutExtension(source));
                int done = 0;
                foreach (var m in messages)
                {
                    if (ShouldCancel()) { result.Cancelled = true; break; }
                    done++;
                    string key = null;
                    try
                    {
                        if (opts.ExcludeFolders.Any(x => string.Equals(x, m.Folder.Path.FirstOrDefault(), StringComparison.OrdinalIgnoreCase))) key = null;
                        else key = GroupKey(m);
                    }
                    catch (Exception ex) when (ex is PstFormatException || ex is ArgumentException || ex is IndexOutOfRangeException)
                    {
                        result.Failed++;
                        Log(string.Format("  Skipping a damaged item in {0}: {1}", string.Join("\\", m.Folder.Path), ex.Message));
                        continue;
                    }
                    if (key == null) { result.Skipped++; }
                    else
                    {
                        if (!groups.TryGetValue(key, out var parts)) groups[key] = parts = new List<Group>();
                        var g = parts.LastOrDefault();
                        if (g == null || (maxBytes > 0 && g.Messages.Count > 0 && skeletonBytes + g.Bytes + m.Bytes > maxBytes))
                        {
                            g = new Group { Key = key, Part = parts.Count + 1 };
                            parts.Add(g);
                            order.Add(g);
                        }
                        g.Messages.Add(m);
                        g.Bytes += m.Bytes;
                    }
                    if (done % 25 == 0 || done == messages.Count) Progress(done, messages.Count, string.Format("Reading items: {0:N0} of {1:N0}", done, messages.Count));
                }

                foreach (var g in order)
                {
                    string name = baseName;
                    if (g.Key.Length > 0) name += "_" + g.Key;
                    if (maxBytes > 0) name += string.Format("_part{0:000}", g.Part);
                    g.File = new SplitFile { Name = name + ".pst", Path = Path.Combine(outDir, name + ".pst"), Items = g.Messages.Count, Bytes = skeletonBytes + g.Bytes };
                }

                if (!preview && !result.Cancelled && order.Count > 0)
                {
                    Directory.CreateDirectory(outDir);
                    foreach (var g in order)
                        if (File.Exists(g.File.Path)) throw new IOException("Output file already exists: " + g.File.Path + ". Choose an empty output folder.");
                    CheckFreeSpace(outDir, order.Sum(g => g.File.Bytes));

                    for (int i = 0; i < order.Count && !result.Cancelled; i++)
                    {
                        var g = order[i];
                        Log("Writing " + g.File.Name);
                        try
                        {
                            g.File.Bytes = WriteOutput(g, i + 1, order.Count);
                        }
                        catch (OperationCanceledException)
                        {
                            try { File.Delete(g.File.Path); } catch { }
                            result.Cancelled = true;
                            break;
                        }
                        catch
                        {
                            try { File.Delete(g.File.Path); } catch { }
                            throw;
                        }
                        result.Files.Add(g.File);
                        result.Copied += g.Messages.Count;
                    }
                }
                else
                {
                    foreach (var g in order) { result.Files.Add(g.File); result.Copied += g.Messages.Count; }
                }
            }

            Log("");
            if (result.Cancelled) Log("Stopped. The files written so far are complete and can be opened.");
            if (preview) Log("Preview only. No files were written.");
            foreach (var f in result.Files)
                Log(string.Format("{0}  {1:N0} items, {2:N1} MB", f.Name, f.Items, f.Bytes / 1048576.0));
            Log(string.Format("Done. {0:N0} items in {1} file(s), {2:N0} skipped, {3:N0} could not be read.", result.Copied, result.Files.Count, result.Skipped, result.Failed));
            return result;
        }

        static void CheckFreeSpace(string dir, long needed)
        {
            try
            {
                var drive = new DriveInfo(Path.GetPathRoot(dir));
                if (drive.IsReady && drive.AvailableFreeSpace < needed + 50L * 1024 * 1024)
                    throw new IOException(string.Format("Not enough free space on {0}. The split files need about {1:N0} MB and {2:N0} MB is free.",
                        drive.Name, needed / 1048576.0, drive.AvailableFreeSpace / 1048576.0));
            }
            catch (ArgumentException) { }
        }

        long TreeBytes(ulong bid)
        {
            if (bid == 0) return 0;
            var set = new HashSet<ulong>();
            pst.CollectTreeBlocks(bid, set);
            long total = 0;
            // Block space and its index entry, plus ~3% for allocation-map pages and the space left
            // where a block can't straddle one, so files stay under the size limit.
            foreach (var b in set) total += (pst.GetBlock(b).Cb + 16 + 63) / 64 * 64 + 28;
            return total + total / 32;
        }

        PropertyContext OpenProps(ulong nid)
        {
            if (!pst.Nodes.TryGetValue(nid, out var n) || n.BidData == 0) return null;
            return PropertyContext.Open(pst, n.BidData, n.BidSub);
        }

        TableContext OpenTable(ulong nid)
        {
            if (!pst.Nodes.TryGetValue(nid, out var n) || n.BidData == 0) return null;
            return TableContext.Open(pst, n.BidData, n.BidSub);
        }

        static ulong EntryIdNid(byte[] entryId) => entryId != null && entryId.Length >= 24 ? BitConverter.ToUInt32(entryId, 20) : 0;

        byte[] recordKey;
        ulong ipmSubtree, wastebasket, finder;

        void LoadFolders()
        {
            var store = OpenProps(Nid.MessageStore) ?? throw new PstFormatException("The PST's store information is missing.");
            recordKey = store.Bytes(PidRecordKey);
            ipmSubtree = EntryIdNid(store.Bytes(PidIpmSubtree));
            wastebasket = EntryIdNid(store.Bytes(PidWastebasket));
            finder = EntryIdNid(store.Bytes(PidFinder));
            if (recordKey == null || recordKey.Length != 16 || ipmSubtree == 0)
                throw new PstFormatException("The PST's store information is incomplete. Run the Inbox Repair Tool (scanpst.exe) on it first.");

            folders = new Dictionary<ulong, FolderInfo>();
            foreach (var n in pst.Nodes.Values)
            {
                uint t = Nid.Type(n.Nid);
                if (t != Nid.TypeNormalFolder && t != Nid.TypeSearchFolder) continue;
                string name = null;
                try { name = OpenProps(n.Nid)?.String(PidDisplayName); } catch (PstFormatException) { }
                folders[n.Nid] = new FolderInfo { Nid = n.Nid, Parent = n.NidParent, Name = name ?? "", IsSearch = t == Nid.TypeSearchFolder };
            }
            foreach (var f in folders.Values)
                if (f.Parent != f.Nid && folders.TryGetValue(f.Parent, out var p)) p.Children.Add(f);

            if (!folders.TryGetValue(ipmSubtree, out var top)) throw new PstFormatException("The PST's top folder is missing.");
            void SetPaths(FolderInfo f, string[] path)
            {
                f.Path = path;
                foreach (var c in f.Children) SetPaths(c, path.Concat(new[] { c.Name }).ToArray());
            }
            SetPaths(top, new string[0]);

            // Folders Outlook refers to by ID (Inbox, Calendar, Drafts, ...) stay in every output.
            var special = new HashSet<ulong> { Nid.RootFolder, ipmSubtree, wastebasket, finder };
            var candidates = pst.Nodes.Values.Where(n => n.Nid == Nid.MessageStore || Nid.Type(n.Nid) == Nid.TypeNormalFolder || Nid.Type(n.Nid) == 0x0B || Nid.Type(n.Nid) == 0x0C);
            foreach (var n in candidates)
            {
                if (n.BidData == 0) continue;
                byte[] data;
                try { data = pst.ReadData(n.BidData); } catch (PstFormatException) { continue; }
                for (int i = 0; i + 24 <= data.Length; i++)
                {
                    int k = 0;
                    while (k < 16 && data[i + k] == recordKey[k]) k++;
                    if (k == 16) special.Add(BitConverter.ToUInt32(data, i + 16));
                }
            }
            foreach (var f in folders.Values) f.Special = f.Path == null || special.Contains(f.Nid);
        }

        /// <summary>Property ID of a numeric named property, from the name-to-ID map ([MS-PST] 2.4.7).</summary>
        ushort? FindNamedProperty(Guid guid, uint lid)
        {
            try
            {
                var map = OpenProps(Nid.NameToIdMap);
                if (map == null) return null;
                var guids = map.Bytes(0x0002) ?? new byte[0];
                var entries = map.Bytes(0x0003) ?? new byte[0];
                int guidIndex = -1;
                var target = guid.ToByteArray();
                for (int i = 0; i + 16 <= guids.Length; i += 16)
                    if (guids.Skip(i).Take(16).SequenceEqual(target)) { guidIndex = i / 16 + 3; break; }
                if (guidIndex < 0) return null;
                for (int i = 0; i + 8 <= entries.Length; i += 8)
                {
                    uint id = BitConverter.ToUInt32(entries, i);
                    ushort g = BitConverter.ToUInt16(entries, i + 4);
                    ushort idx = BitConverter.ToUInt16(entries, i + 6);
                    if ((g & 1) == 0 && (g >> 1) == guidIndex && id == lid) return (ushort)(0x8000 + idx);
                }
            }
            catch (PstFormatException) { }
            return null;
        }

        // ---- grouping --------------------------------------------------------

        int folderKeyDepth;

        /// <summary>
        /// Folder mode splits on the top-level folders. When everything sits under a single folder
        /// (e.g. one old mailbox folder), it splits on the first level that has more than one folder.
        /// </summary>
        void SetFolderKeys(List<MessageInfo> messages)
        {
            var paths = messages.Select(m => m.Folder.Path)
                .Where(p => p.Length > 0 && !opts.ExcludeFolders.Any(x => string.Equals(x, p[0], StringComparison.OrdinalIgnoreCase)))
                .Distinct().ToList();
            folderKeyDepth = 0;
            while (paths.Where(p => p.Length > folderKeyDepth).Select(p => p[folderKeyDepth]).Distinct().Count() == 1 &&
                   paths.Any(p => p.Length > folderKeyDepth + 1))
                folderKeyDepth++;
        }

        string GroupKey(MessageInfo m)
        {
            switch (opts.Mode)
            {
                case SplitMode.Size:
                    return "";
                case SplitMode.Folder:
                    return m.Folder.Path.Length <= folderKeyDepth ? SafeName(m.Folder.Path.LastOrDefault() ?? "Top of folders") : SafeName(m.Folder.Path[folderKeyDepth]);
                case SplitMode.Year:
                {
                    var d = ItemDate(m);
                    return d.HasValue ? d.Value.ToString("yyyy") : "NoDate";
                }
                case SplitMode.Month:
                {
                    var d = ItemDate(m);
                    return d.HasValue ? d.Value.ToString("yyyy-MM") : "NoDate";
                }
                case SplitMode.Date:
                {
                    var d = ItemDate(m);
                    bool inside = d.HasValue && d.Value >= opts.StartDate.Date && d.Value < opts.EndDate.Date.AddDays(1);
                    if (inside) return string.Format("{0:yyyy-MM-dd}_to_{1:yyyy-MM-dd}", opts.StartDate, opts.EndDate);
                    return opts.IncludeOther ? "Other" : null;
                }
                case SplitMode.Sender:
                {
                    var addrs = Addresses(m);
                    if (opts.EmailAddresses.Count > 0)
                    {
                        foreach (var a in addrs)
                            if (opts.EmailAddresses.Contains(a)) return SafeName(a);
                        return opts.IncludeOther ? "Other" : null;
                    }
                    return addrs.Count > 0 ? SafeName(addrs[0]) : "Unknown";
                }
            }
            return null;
        }

        DateTime? ItemDate(MessageInfo m)
        {
            var pc = OpenProps(m.Nid);
            if (pc == null) return null;
            DateTime? d = pc.Time(PidDeliveryTime) ?? pc.Time(PidSubmitTime);
            if (!d.HasValue && appointmentStartProp.HasValue) d = pc.Time(appointmentStartProp.Value);
            return d ?? pc.Time(PidCreationTime);
        }

        List<string> Addresses(MessageInfo m)
        {
            var list = new List<string>();
            var pc = OpenProps(m.Nid);
            if (pc == null) return list;
            if (opts.MatchOn != MatchOn.Recipients)
            {
                string a = pc.String(PidSenderSmtp);
                if (string.IsNullOrEmpty(a) && string.Equals(pc.String(PidSenderAddrType), "SMTP", StringComparison.OrdinalIgnoreCase)) a = pc.String(PidSenderEmail);
                if (string.IsNullOrEmpty(a)) a = pc.String(PidSentRepSmtp);
                if (string.IsNullOrEmpty(a)) { var e = pc.String(PidSentRepEmail); if (e != null && e.Contains("@")) a = e; }
                if (string.IsNullOrEmpty(a)) { var e = pc.String(PidSenderEmail); if (e != null && e.Contains("@")) a = e; }
                if (string.IsNullOrEmpty(a)) a = pc.String(PidSenderName);
                if (!string.IsNullOrWhiteSpace(a)) list.Add(a.Trim().ToLowerInvariant());
            }
            if (opts.MatchOn != MatchOn.Sender && pc.Heap.Subnodes.TryGetValue(Nid.RecipientTable, out var sn))
            {
                var rt = TableContext.Open(pst, sn);
                if (rt != null)
                    foreach (var row in rt.Rows)
                    {
                        string a = rt.CellString(row, TagSmtp);
                        if (string.IsNullOrEmpty(a)) { var e = rt.CellString(row, TagEmail); if (e != null && e.Contains("@")) a = e; }
                        if (!string.IsNullOrWhiteSpace(a)) list.Add(a.Trim().ToLowerInvariant());
                    }
            }
            return list;
        }

        // ---- writing ---------------------------------------------------------

        long WriteOutput(Group g, int fileNumber, int fileCount)
        {
            var keepMessages = new HashSet<ulong>(g.Messages.Select(m => m.Nid));
            var unread = new HashSet<ulong>(g.Messages.Where(m => m.Unread).Select(m => m.Nid));

            // Folders to keep: special ones, plus every folder holding a kept message and its parents.
            var keepFolders = new HashSet<ulong>(folders.Values.Where(f => f.Special).Select(f => f.Nid));
            foreach (var m in g.Messages)
                for (var f = m.Folder; f != null && keepFolders.Add(f.Nid);)
                    f = folders.TryGetValue(f.Parent, out var p) && p != f ? p : null;
            foreach (var f in folders.Values.Where(f => keepFolders.Contains(f.Nid)).ToList())
                for (var p = f; folders.TryGetValue(p.Parent, out var parent) && parent != p; p = parent) keepFolders.Add(parent.Nid);
            var droppedIndexes = new HashSet<uint>(folders.Values.Where(f => !keepFolders.Contains(f.Nid)).Select(f => Nid.Index(f.Nid)));
            var droppedFolders = new HashSet<ulong>(folders.Values.Where(f => !keepFolders.Contains(f.Nid)).Select(f => f.Nid));

            // A new store ID, so Outlook can open this file alongside the original and the other parts.
            var newKey = Guid.NewGuid().ToByteArray();
            Func<byte[], byte[]> swapKey = v => LtpWriter.ReplaceBytes(new[] { v }, recordKey, newKey)?[0] ?? v;
            var w = new PstWriter(pst);

            // Rows kept per folder table, for counts in folder properties and hierarchy tables.
            var counts = new Dictionary<ulong, KeyValuePair<int, int>>();
            var hasChildren = new Dictionary<ulong, bool>();
            foreach (var f in folders.Values)
                hasChildren[f.Nid] = f.Children.Any(c => keepFolders.Contains(c.Nid) && c.Nid != f.Nid);

            var tableUpdates = new Dictionary<ulong, KeyValuePair<ulong, ulong>>();
            foreach (var f in folders.Values.Where(f => keepFolders.Contains(f.Nid)))
            {
                ulong ctNid = Nid.Make(f.IsSearch ? Nid.TypeSearchContentsTable : Nid.TypeContentsTable, Nid.Index(f.Nid));
                var ct = OpenTable(ctNid);
                if (ct == null) continue;
                var rows = ct.Rows.Where(r => keepMessages.Contains(TableContext.RowId(r))).ToList();
                counts[f.Nid] = new KeyValuePair<int, int>(rows.Count, rows.Count(r => unread.Contains(TableContext.RowId(r))));
                if (rows.Count != ct.Rows.Count) tableUpdates[ctNid] = LtpWriter.RebuildTable(w, ct, rows, swapKey);
            }

            foreach (var f in folders.Values.Where(f => keepFolders.Contains(f.Nid) && !f.IsSearch))
            {
                ulong htNid = Nid.Make(Nid.TypeHierarchyTable, Nid.Index(f.Nid));
                var ht = OpenTable(htNid);
                if (ht == null) continue;
                int cCount = ht.ColumnIndex(TagContentCount), cUnread = ht.ColumnIndex(TagContentUnread), cSub = ht.ColumnIndex(TagSubfolders);
                var rows = new List<byte[]>();
                bool changed = false;
                foreach (var r in ht.Rows)
                {
                    ulong child = TableContext.RowId(r);
                    if (!keepFolders.Contains(child)) { changed = true; continue; }
                    var row = (byte[])r.Clone();
                    if (counts.TryGetValue(child, out var c))
                    {
                        changed |= SetCell(ht, row, cCount, c.Key);
                        changed |= SetCell(ht, row, cUnread, c.Value);
                    }
                    if (hasChildren.TryGetValue(child, out bool hc)) changed |= SetCell(ht, row, cSub, hc ? 1 : 0);
                    rows.Add(row);
                }
                if (changed) tableUpdates[htNid] = LtpWriter.RebuildTable(w, ht, rows, swapKey);
            }

            foreach (var n in pst.Nodes.Values)
            {
                uint type = Nid.Type(n.Nid);
                if (type == Nid.TypeNormalMessage && !keepMessages.Contains(n.Nid)) continue;
                if (droppedFolders.Contains(n.NidParent)) continue;
                if ((type == Nid.TypeNormalFolder || type == Nid.TypeSearchFolder || type == Nid.TypeHierarchyTable ||
                     type == Nid.TypeContentsTable || type == Nid.TypeAssocContentsTable || type == Nid.TypeSearchContentsTable ||
                     type == 0x06 || type == 0x07) && droppedIndexes.Contains(Nid.Index(n.Nid)))
                    continue;

                if (tableUpdates.TryGetValue(n.Nid, out var t))
                {
                    w.AddNode(n.Nid, t.Key, t.Value, n.NidParent);
                    continue;
                }
                if (type == Nid.TypeNormalMessage)
                {
                    w.AddNode(n.Nid, w.CopySource(n.BidData), w.CopySource(n.BidSub), n.NidParent);
                    continue;
                }
                if (type == 0x06 && n.BidData != 0)
                {
                    // Search update queue: pending work may refer to items not in this file. Outlook rebuilds it.
                    w.AddNode(n.Nid, 0, 0, n.NidParent);
                    continue;
                }

                // Other nodes: copy, changing folder counts and the store ID where they appear.
                ulong bidData = n.BidData;
                if (n.BidData != 0 && type != Nid.TypeAssocMessage)
                {
                    List<byte[]> blocks = null;
                    var original = pst.DataBlockIds(n.BidData).Select(pst.ReadBlock).ToList();
                    if ((type == Nid.TypeNormalFolder || type == Nid.TypeSearchFolder) && TryOpenHeap(n) is HeapNode heap && heap.ClientSig == 0xBC)
                    {
                        var values = new Dictionary<ushort, uint>();
                        if (counts.TryGetValue(n.Nid, out var c)) { values[PidContentCount] = (uint)c.Key; values[PidContentUnread] = (uint)c.Value; }
                        if (!folders[n.Nid].IsSearch) values[PidSubfolders] = hasChildren[n.Nid] ? 1u : 0u;
                        blocks = LtpWriter.PatchProperties(heap, values);
                    }
                    var replaced = LtpWriter.ReplaceBytes(blocks ?? original, recordKey, newKey);
                    if (replaced != null) blocks = replaced;
                    if (blocks != null) bidData = w.NewDataTree(blocks);
                }
                w.AddNode(n.Nid, bidData == n.BidData ? w.CopySource(n.BidData) : bidData, w.CopySource(n.BidSub), n.NidParent);
            }

            string label = fileCount > 1 ? string.Format("Writing file {0} of {1}", fileNumber, fileCount) : "Writing";
            int lastPercent = -1;
            w.Write(g.File.Path, ShouldCancel, (doneBytes, totalBytes) =>
            {
                int pct = totalBytes == 0 ? 100 : (int)(doneBytes * 100 / totalBytes);
                if (pct != lastPercent) { lastPercent = pct; Progress(pct, 100, string.Format("{0}: {1} ({2}%)", label, g.File.Name, pct)); }
            });
            return w.BytesWritten;
        }

        HeapNode TryOpenHeap(NodeEntry n)
        {
            try
            {
                var heap = new HeapNode(pst, n.BidData, n.BidSub);
                return heap;
            }
            catch (PstFormatException) { return null; }
        }

        static bool SetCell(TableContext t, byte[] row, int col, int value)
        {
            if (col < 0) return false;
            var c = t.Columns[col];
            var bytes = c.Cb == 1 ? new[] { (byte)value } : BitConverter.GetBytes(value);
            bool same = t.CellExists(row, c);
            for (int i = 0; i < c.Cb && same; i++) same = row[c.Ib + i] == bytes[i];
            if (same) return false;
            Buffer.BlockCopy(bytes, 0, row, c.Ib, Math.Min(bytes.Length, c.Cb));
            row[t.Rgib[2] + c.Bit / 8] |= (byte)(0x80 >> (c.Bit % 8));
            return true;
        }

        public static string SafeName(string name)
        {
            var bad = new HashSet<char>(Path.GetInvalidFileNameChars());
            var clean = new string((name ?? "").Select(c => bad.Contains(c) ? '_' : c).ToArray()).Trim().TrimEnd('.');
            if (clean.Length > 80) clean = clean.Substring(0, 80);
            return clean.Length == 0 ? "unnamed" : clean;
        }
    }
}
