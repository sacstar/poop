using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace PstSplitter
{
    public enum SplitMode { Size, Year, Month, Date, Folder, Sender }
    public enum MatchOn { Sender, Recipients, Both }

    public class SplitOptions
    {
        public string Source;
        public string Output;
        public SplitMode Mode = SplitMode.Size;
        public int MaxSizeMB;
        public DateTime StartDate;
        public DateTime EndDate;
        public List<string> EmailAddresses = new List<string>();
        public MatchOn MatchOn = MatchOn.Sender;
        public bool IncludeOther;
        public List<string> ExcludeFolders = new List<string>();

        public static readonly string[] DefaultExcludes =
        {
            "Deleted Items", "Junk Email", "Junk E-mail", "Search Folders", "Sync Issues",
            "Conflicts", "Local Failures", "Server Failures"
        };

        public static List<string> ParseAddresses(string text)
        {
            return Regex.Split(text ?? "", @"[,;\s]+")
                .Select(a => a.Trim().ToLowerInvariant())
                .Where(a => a.Length > 0)
                .Distinct()
                .ToList();
        }

        /// <summary>Returns a message describing what's wrong, or null when the options are usable.</summary>
        public string Validate()
        {
            if (string.IsNullOrWhiteSpace(Source)) return "Choose the PST file to split.";
            if (!File.Exists(Source)) return "The PST file was not found.";
            if (string.IsNullOrWhiteSpace(Output)) return "Choose an output folder.";
            if (Mode == SplitMode.Size && MaxSizeMB <= 0) return "Enter a maximum size in MB.";
            if (Mode == SplitMode.Date && EndDate < StartDate) return "The end date is before the start date.";
            return null;
        }
    }

    public class SplitFile
    {
        public string Name;
        public string Path;
        public int Items;
        public long Bytes;
    }

    public class SplitResult
    {
        public List<SplitFile> Files = new List<SplitFile>();
        public int Copied, Skipped, Failed;
        public bool Cancelled, Preview;
        public string Output;
    }

    /// <summary>
    /// Splits a PST by driving Outlook through late-bound COM. Items are copied
    /// into new PST files with the same folder structure; the source is only read.
    /// </summary>
    public class Splitter
    {
        const int OlStoreUnicode = 3;
        const int OlExchangeUserAddressEntry = 0;
        const int OlExchangeRemoteUserAddressEntry = 5;
        // Folder.DefaultItemType -> folder type for Folders.Add
        static readonly Dictionary<int, int> FolderTypeForItemType = new Dictionary<int, int>
        {
            { 0, 6 }, { 1, 9 }, { 2, 10 }, { 3, 13 }, { 4, 11 }, { 5, 12 }
        };

        public Action<string> Log = _ => { };
        public Action<int, int, string> Progress = (done, total, folder) => { };
        public Func<bool> ShouldCancel = () => false;
        public Func<object> CreateOutlook = () =>
        {
            var type = Type.GetTypeFromProgID("Outlook.Application");
            if (type == null) throw new InvalidOperationException("Microsoft Outlook is not installed on this PC. PST Splitter needs Outlook to read and write PST files.");
            return Activator.CreateInstance(type);
        };

        class SourceFolder
        {
            public dynamic Folder;
            public string[] Path;
            public int ItemType;
        }

        class Group
        {
            public string Key;
            public int Part;
            public SplitFile File;
            public dynamic Store;
            public dynamic Root;
            public Dictionary<string, dynamic> Folders = new Dictionary<string, dynamic>(StringComparer.OrdinalIgnoreCase);
        }

        SplitOptions opts;
        bool dryRun;
        dynamic ns;
        string outDir, baseName;
        long maxBytes;
        readonly Dictionary<string, Group> groups = new Dictionary<string, Group>();
        readonly List<Group> opened = new List<Group>();

        public SplitResult Run(SplitOptions options, bool preview)
        {
            opts = options;
            dryRun = preview;
            groups.Clear();
            opened.Clear();

            var source = System.IO.Path.GetFullPath(opts.Source);
            if (!File.Exists(source)) throw new FileNotFoundException("Source PST not found: " + source);
            outDir = System.IO.Path.GetFullPath(opts.Output);
            if (!dryRun) Directory.CreateDirectory(outDir);
            maxBytes = (long)opts.MaxSizeMB * 1024 * 1024;
            baseName = SafeName(System.IO.Path.GetFileNameWithoutExtension(source));

            var result = new SplitResult { Preview = dryRun, Output = outDir };
            Log("Starting Outlook...");
            dynamic outlook = CreateOutlook();
            ns = outlook.GetNamespace("MAPI");

            dynamic srcStore = FindStore(source);
            bool weOpenedSource = false;
            if (srcStore == null)
            {
                Log("Opening " + source);
                ns.AddStoreEx(source, OlStoreUnicode);
                srcStore = FindStore(source);
                weOpenedSource = true;
            }
            if (srcStore == null) throw new InvalidOperationException("Outlook could not open " + source);
            string srcStoreId = srcStore.StoreID;

            var folders = new List<SourceFolder>();
            dynamic srcRoot = srcStore.GetRootFolder();
            CollectFolders(srcRoot, new string[0], folders);
            int total = 0;
            foreach (var f in folders)
            {
                try { total += (int)f.Folder.Items.Count; } catch { }
            }
            Log(string.Format("Found {0} folders and {1:N0} items", folders.Count, total));

            int done = 0;
            try
            {
                foreach (var f in folders)
                {
                    if (result.Cancelled) break;
                    string pathText = string.Join("\\", f.Path);
                    // Collect IDs first: copying adds items to the folder while we work.
                    var ids = GetEntryIds(f.Folder);
                    if (ids.Count == 0) continue;
                    Log(string.Format("{0}: {1:N0} items", pathText, ids.Count));

                    foreach (var id in ids)
                    {
                        if (ShouldCancel()) { result.Cancelled = true; break; }
                        done++;
                        dynamic item = null, copy = null, moved = null;
                        try
                        {
                            item = ns.GetItemFromID(id, srcStoreId);
                            string key = GetGroupKey(item, f.Path);
                            if (key == null) { result.Skipped++; continue; }
                            long size = 0;
                            try { size = (long)item.Size; } catch { }
                            var g = GetTarget(key, size);

                            if (!dryRun)
                            {
                                if (!g.Folders.TryGetValue(pathText, out dynamic dest))
                                {
                                    dest = g.Root;
                                    foreach (var part in f.Path) dest = GetOrAddFolder(dest, part, f.ItemType);
                                    g.Folders[pathText] = dest;
                                }
                                copy = item.Copy();
                                moved = copy.Move(dest);
                            }
                            g.File.Bytes += size;
                            g.File.Items++;
                            result.Copied++;
                        }
                        catch (Exception ex)
                        {
                            result.Failed++;
                            // Don't leave a stray copy in the source if the move failed.
                            if (copy != null && moved == null) { try { copy.Delete(); } catch { } }
                            Log(string.Format("  Could not copy an item in {0}: {1}", pathText, ex.Message));
                        }
                        finally
                        {
                            Release(moved); Release(copy); Release(item);
                            Progress(done, total, pathText);
                            if (done % 200 == 0) { GC.Collect(); GC.WaitForPendingFinalizers(); }
                        }
                    }
                }
            }
            finally
            {
                foreach (var g in opened)
                {
                    if (g.Root == null) continue;
                    try { ns.RemoveStore(g.Root); } catch { }
                    foreach (var folder in g.Folders.Values) Release(folder);
                    Release(g.Root); Release(g.Store);
                }
                if (weOpenedSource) { try { ns.RemoveStore(srcRoot); } catch { } }
                foreach (var f in folders) Release(f.Folder);
                Release(srcRoot); Release(srcStore); Release(ns); Release(outlook);
                ns = null;
                GC.Collect(); GC.WaitForPendingFinalizers();
            }

            Log("");
            if (result.Cancelled) Log("Stopped. The files written so far are complete and can be opened.");
            if (dryRun) Log("Preview only. No files were written.");
            foreach (var g in opened)
            {
                result.Files.Add(g.File);
                Log(string.Format("{0}  {1:N0} items, {2:N1} MB", g.File.Name, g.File.Items, g.File.Bytes / 1048576.0));
            }
            Log(string.Format("Done. Copied {0:N0}, skipped {1:N0}, failed {2:N0}.", result.Copied, result.Skipped, result.Failed));
            return result;
        }

        Group GetTarget(string key, long size)
        {
            groups.TryGetValue(key, out var g);
            if (g != null && maxBytes > 0 && g.File.Bytes > 0 && g.File.Bytes + size > maxBytes)
                g = NewGroup(key, g.Part + 1);
            return g ?? NewGroup(key, 1);
        }

        Group NewGroup(string key, int part)
        {
            string name = baseName;
            if (key.Length > 0) name += "_" + key;
            if (maxBytes > 0) name += string.Format("_part{0:000}", part);
            string path = System.IO.Path.Combine(outDir, name + ".pst");
            var g = new Group { Key = key, Part = part, File = new SplitFile { Name = name + ".pst", Path = path } };
            if (!dryRun)
            {
                if (File.Exists(path)) throw new IOException("Output file already exists: " + path + ". Choose an empty output folder.");
                Log("Creating " + path);
                ns.AddStoreEx(path, OlStoreUnicode);
                g.Store = FindStore(path);
                if (g.Store == null) throw new InvalidOperationException("Outlook could not create " + path);
                g.Root = g.Store.GetRootFolder();
                try { g.Root.Name = name; } catch { }
            }
            groups[key] = g;
            opened.Add(g);
            return g;
        }

        string GetGroupKey(dynamic item, string[] folderPath)
        {
            switch (opts.Mode)
            {
                case SplitMode.Size:
                    return "";
                case SplitMode.Folder:
                    return SafeName(folderPath[0]);
                case SplitMode.Year:
                {
                    DateTime? d = GetItemDate(item);
                    return d.HasValue ? d.Value.ToString("yyyy") : "NoDate";
                }
                case SplitMode.Month:
                {
                    DateTime? d = GetItemDate(item);
                    return d.HasValue ? d.Value.ToString("yyyy-MM") : "NoDate";
                }
                case SplitMode.Date:
                {
                    DateTime? d = GetItemDate(item);
                    bool inside = d.HasValue && d.Value >= opts.StartDate.Date && d.Value < opts.EndDate.Date.AddDays(1);
                    if (inside) return string.Format("{0:yyyy-MM-dd}_to_{1:yyyy-MM-dd}", opts.StartDate, opts.EndDate);
                    return opts.IncludeOther ? "Other" : null;
                }
                case SplitMode.Sender:
                {
                    List<string> addrs = GetItemAddresses(item);
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

        static DateTime? GetItemDate(dynamic item)
        {
            foreach (var prop in new[] { "ReceivedTime", "SentOn", "Start", "CreationTime" })
            {
                try
                {
                    object v = null;
                    switch (prop)
                    {
                        case "ReceivedTime": v = item.ReceivedTime; break;
                        case "SentOn": v = item.SentOn; break;
                        case "Start": v = item.Start; break;
                        case "CreationTime": v = item.CreationTime; break;
                    }
                    // Outlook uses 1/1/4501 for "none"
                    if (v is DateTime d && d.Year < 4500 && d.Year > 1601) return d;
                }
                catch { }
            }
            return null;
        }

        List<string> GetItemAddresses(dynamic item)
        {
            var list = new List<string>();
            if (opts.MatchOn != MatchOn.Recipients)
            {
                string fallback = null;
                try { fallback = item.SenderEmailAddress; } catch { }
                dynamic sender = null;
                try { sender = item.Sender; } catch { }
                string a = GetSmtpAddress(sender, fallback);
                Release(sender);
                if (a != null) list.Add(a);
            }
            if (opts.MatchOn != MatchOn.Sender)
            {
                try
                {
                    dynamic recips = item.Recipients;
                    int count = recips.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        dynamic r = recips.Item(i);
                        string fallback = null;
                        try { fallback = r.Address; } catch { }
                        dynamic entry = null;
                        try { entry = r.AddressEntry; } catch { }
                        string a = GetSmtpAddress(entry, fallback);
                        if (a != null) list.Add(a);
                        Release(entry); Release(r);
                    }
                    Release(recips);
                }
                catch { }
            }
            return list;
        }

        static string GetSmtpAddress(dynamic entry, string fallback)
        {
            try
            {
                if (entry != null)
                {
                    int t = entry.AddressEntryUserType;
                    if (t == OlExchangeUserAddressEntry || t == OlExchangeRemoteUserAddressEntry)
                    {
                        dynamic ex = entry.GetExchangeUser();
                        if (ex != null)
                        {
                            string smtp = ex.PrimarySmtpAddress;
                            Release(ex);
                            if (!string.IsNullOrEmpty(smtp)) return smtp.ToLowerInvariant();
                        }
                    }
                    string addr = entry.Address;
                    if (!string.IsNullOrEmpty(addr) && addr.Contains("@")) return addr.ToLowerInvariant();
                }
            }
            catch { }
            return string.IsNullOrEmpty(fallback) ? null : fallback.ToLowerInvariant();
        }

        dynamic FindStore(string path)
        {
            dynamic stores = ns.Stores;
            int count = stores.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic s = stores.Item(i);
                try
                {
                    string fp = s.FilePath;
                    if (!string.IsNullOrEmpty(fp) && string.Equals(System.IO.Path.GetFullPath(fp), path, StringComparison.OrdinalIgnoreCase))
                        return s;
                }
                catch { }
                Release(s);
            }
            return null;
        }

        static dynamic GetOrAddFolder(dynamic parent, string name, int itemType)
        {
            dynamic children = parent.Folders;
            int count = children.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic f = children.Item(i);
                if (string.Equals((string)f.Name, name, StringComparison.OrdinalIgnoreCase)) return f;
                Release(f);
            }
            int type;
            if (!FolderTypeForItemType.TryGetValue(itemType, out type)) type = 6;
            return children.Add(name, type);
        }

        // Every folder below the root with its path, e.g. { "Inbox", "Clients" }.
        void CollectFolders(dynamic folder, string[] path, List<SourceFolder> result)
        {
            dynamic children = folder.Folders;
            int count = children.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic sub = children.Item(i);
                string name = sub.Name;
                if (path.Length == 0 && opts.ExcludeFolders.Any(x => string.Equals(x, name, StringComparison.OrdinalIgnoreCase)))
                {
                    Release(sub);
                    continue;
                }
                var subPath = path.Concat(new[] { name }).ToArray();
                int itemType = 0;
                try { itemType = sub.DefaultItemType; } catch { }
                result.Add(new SourceFolder { Folder = sub, Path = subPath, ItemType = itemType });
                CollectFolders(sub, subPath, result);
            }
        }

        static List<string> GetEntryIds(dynamic folder)
        {
            var ids = new List<string>();
            // A table reads IDs without opening each item, which is much faster.
            try
            {
                dynamic table = folder.GetTable();
                while (!(bool)table.EndOfTable)
                {
                    dynamic row = table.GetNextRow();
                    ids.Add((string)row["EntryID"]);
                    Release(row);
                }
                Release(table);
                return ids;
            }
            catch { ids.Clear(); }

            dynamic items = folder.Items;
            dynamic it = items.GetFirst();
            while (it != null)
            {
                try { ids.Add((string)it.EntryID); } catch { }
                Release(it);
                it = items.GetNext();
            }
            Release(items);
            return ids;
        }

        public static string SafeName(string name)
        {
            var bad = new HashSet<char>(System.IO.Path.GetInvalidFileNameChars());
            var clean = new string((name ?? "").Select(c => bad.Contains(c) ? '_' : c).ToArray()).Trim().TrimEnd('.');
            if (clean.Length > 80) clean = clean.Substring(0, 80);
            return clean.Length == 0 ? "unnamed" : clean;
        }

        static void Release(object o)
        {
            if (o != null && Marshal.IsComObject(o)) Marshal.ReleaseComObject(o);
        }
    }
}
