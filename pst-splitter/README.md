# PST Splitter

Splits a large Outlook `.pst` file into smaller `.pst` files, like vMail PST
Splitter. It runs on Windows and uses your installed Microsoft Outlook to read
and write the files, so every output file opens in Outlook as normal.

The source PST is only read. Items are copied into the new files, and the
folder structure (Inbox, Inbox\Clients, Sent Items, Calendar, Contacts...)
is kept in each one.

## Split modes

| Mode | Result |
|---|---|
| Size | A new file each time one reaches the size limit (`archive_part001.pst`, `_part002`...) |
| Year | One file per year (`archive_2023.pst`) |
| Month | One file per month (`archive_2023-07.pst`) |
| Date range | One file with items between two dates; optionally an `_Other` file for the rest |
| Top-level folder | One file per folder under the root (`archive_Inbox.pst`, `archive_Calendar.pst`) |
| Email address | One file per sender, or only the addresses you list (matched on sender, recipients, or both) |

The size limit can be added to any mode. A group that grows past it rolls over
to `_part002`, and so on.

Deleted Items, Junk Email and Outlook's sync/conflict folders are skipped by
default.

## Requirements

- Windows with Microsoft Outlook 2010 or newer (32- or 64-bit)
- Windows PowerShell 5.1 (built into Windows 10/11)

## Using the window

1. Copy `PstSplitter.ps1` and `PstSplitter.bat` to the same folder.
2. Double-click `PstSplitter.bat`.
3. Choose the PST file, the output folder and the split mode.
4. Click **Preview** to see how many items and MB would go to each file without
   writing anything. Then click **Split**.

Outlook opens in the background while it runs. The new files are not left
attached to your Outlook profile. Use File > Open > Outlook Data File to view one.

## Command line

```powershell
.\PstSplitter.ps1 -Source C:\Mail\archive.pst -Output C:\Mail\split -Mode Size -MaxSizeMB 2048
.\PstSplitter.ps1 -Source C:\Mail\archive.pst -Mode Year
.\PstSplitter.ps1 -Source C:\Mail\archive.pst -Mode Year -MaxSizeMB 4096
.\PstSplitter.ps1 -Source C:\Mail\archive.pst -Mode Date -StartDate 2023-01-01 -EndDate 2023-12-31 -IncludeOther
.\PstSplitter.ps1 -Source C:\Mail\archive.pst -Mode Folder
.\PstSplitter.ps1 -Source C:\Mail\archive.pst -Mode Sender -EmailAddresses bob@example.com,amy@example.com -MatchOn Both -IncludeOther
.\PstSplitter.ps1 -Source C:\Mail\archive.pst -Mode Month -WhatIf   # preview only
```

Without `-Output`, files go to a folder named `<source>_split` next to the source.
Use `-ExcludeFolders @()` to include Deleted Items and Junk Email.

## Notes

- **Back up the PST first.** Outlook copies each item inside the source folder
  and then moves the copy out, so the source is briefly written to. It ends up
  unchanged, but a crash mid-run could leave an extra copy behind.
- The size limit uses Outlook's item sizes. A real PST file is a little larger
  than the total of its items, so set the limit about 10% below the size you need.
- Splitting by email address reads sender and recipient addresses. With old or
  missing antivirus, Outlook may ask "A program is trying to access e-mail
  address information". Allow it for the run.
- A password-protected PST asks for its password when it opens.
- Large files take time: roughly 20 to 100 items a second, depending on the PC.
