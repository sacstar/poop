using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
}
