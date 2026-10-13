# PST Splitter

Splits a large Outlook `.pst` file into smaller `.pst` files, like vMail PST
Splitter. **Outlook is not needed.** PST Splitter reads and writes PST files
itself, so it works on any Windows 10 or 11 PC.

The original PST is only read, never changed. Messages are copied into the
new files byte for byte, with their attachments, and keep their folders.

## Split modes

| Mode | Result |
|---|---|
| Size | A new file each time one reaches the size limit (`archive_part001.pst`, `_part002`...) |
| Year | One file per year (`archive_2023.pst`) |
| Month | One file per month (`archive_2023-07.pst`) |
| Date range | One file with items between two dates; optionally an `_Other` file for the rest |
| Folder | One file per main folder (`archive_Inbox.pst`, `archive_Calendar.pst`). If everything sits inside one folder, it splits on that folder's subfolders |
| Email address | One file per sender, or only the addresses you list (matched on sender, recipients, or both) |

The size limit can be added to any mode. A group that grows past it rolls over
to `_part002`, and so on. A single message larger than the limit gets a file
of its own.

Deleted Items, Junk Email and Outlook's sync/conflict folders are skipped by
default.

## Requirements

- Windows 10 or 11. Nothing to install: `PstSplitter.exe` uses .NET Framework 4.8,
  which is built into Windows.
- A Unicode PST (any PST made by Outlook 2003 or later). Old ANSI PSTs from
  Outlook 97-2002 and OST files are not supported.

## Using the program

1. Copy `PstSplitter.exe` anywhere, such as the desktop, and double-click it.
   The exe isn't code-signed, so the first time Windows may show "Windows
   protected your PC". Click **More info**, then **Run anyway**.
2. **Files:** click Browse or drag a `.pst` onto the window. The output folder
   is filled in for you (`<name>_split` next to the PST). Change it if you like.
3. **Split by:** pick Size, Year, Month, Date range, Folder or Email address.
   The options for that mode appear next to it. Tick **Limit each file to** to
   cap the size of every file, whatever the mode.
4. Click **Preview**. The Results tab lists every file that would be made, with
   its item count and estimated size. Nothing is written.
5. Click **Split**. The status line shows each step (reading the index, scanning
   folders, measuring items, writing each file) with a progress bar, and a clock
   shows the elapsed time, so you can see it is working. Click **Stop**
   to end early; files finished before that are complete and can be opened.
6. When it finishes, click **Open output folder**, or double-click a file in
   Results to show it in Explorer.

Open a split file in Outlook with File > Open & Export > Open Outlook Data File.

## What's in each split file

- The messages for that file, in the same folders as in the original.
- Outlook's standard folders (Inbox, Sent Items, Calendar, Contacts and so on),
  even when empty, plus search folders. Other folders appear only when they
  hold messages for that file.
- Folder item counts are updated to match.
- Each file gets its own store ID, so you can open several split files and the
  original in Outlook at the same time.
- A password-protected PST keeps the same password in every split file.

## Notes

- If Outlook has the PST open, close Outlook (or remove the file from Outlook)
  first. PST Splitter says so if the file is in use.
- PST Splitter checks there is enough free disk space before writing.
- If the PST is damaged, run the Inbox Repair Tool (`scanpst.exe`) on a copy
  first. Items that can't be read are listed in the Log tab and skipped.

## How it works

PST Splitter implements Microsoft's published PST file format
([MS-PST](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-pst/)).
Each split file is rebuilt from the original's own structure. The store,
folders, views and named properties are copied from it. Only the chosen messages
are kept, and the folder tables and counts are rebuilt to match. Because
everything else comes from a file Outlook made, the result is laid out the way
Outlook expects.

The code is in `app/`:

- `app/Pst/PstFile.cs`: reads the PST's node and block indexes and blocks
- `app/Pst/Ltp.cs`: reads heaps, property lists and tables
- `app/Pst/PstWriter.cs`: writes a new PST (blocks, B-tree indexes, allocation maps, header)
- `app/Pst/LtpWriter.cs`: rebuilds tables and updates folder properties
- `app/Splitter.cs`: groups messages by mode and writes each split file
- `app/MainForm.cs`: the window

The encoding tables and the allocation-map layout follow Microsoft's open-source
[outlook-pst](https://github.com/microsoft/outlook-pst-rs) library (MIT licence).

### How it was tested

Split files were checked three ways, in every split mode, on a real Outlook PST
(an Enron mailbox) and on larger test PSTs built from it (up to 4,771 items in
one folder, and 60 MB files):

- Every message was read back with the independent libpff library and matched
  to the original: subject, dates, text and HTML body, headers, recipients and
  attachment contents.
- Every folder, table and message was opened with Microsoft's outlook-pst
  library.
- Each file's allocation maps were compared with the ones Microsoft's library
  rebuilds from scratch. They were identical.

## Building the exe

On any PC with the .NET 8 SDK (Windows or Linux):

```
cd app
dotnet build -c Release
```

The exe is written to `app/bin/Release/net48/PstSplitter.exe`. Copy it over
`PstSplitter.exe` here when you change the code.
