<#
.SYNOPSIS
    Splits a large Outlook PST file into smaller PST files.

.DESCRIPTION
    Uses Microsoft Outlook (through COM) to read the source PST and write new
    PST files, so the output opens in Outlook like any other data file.
    The folder structure of the source is kept in every output file.

    Split modes:
      Size    - start a new PST each time the current one reaches -MaxSizeMB
      Year    - one PST per year (by received/sent date)
      Month   - one PST per month
      Date    - one PST with items between -StartDate and -EndDate
      Folder  - one PST per top-level folder (Inbox, Sent Items, Calendar, ...)
      Sender  - one PST per sender email address (or only -EmailAddresses)

    -MaxSizeMB also works with every other mode: a group that grows past the
    limit rolls over to _part2, _part3, and so on.

    The source PST is only read. Items are copied, never moved or deleted.

    Run with no parameters to open the window.

.EXAMPLE
    .\PstSplitter.ps1 -Source C:\Mail\archive.pst -Output C:\Mail\split -Mode Size -MaxSizeMB 2048

.EXAMPLE
    .\PstSplitter.ps1 -Source C:\Mail\archive.pst -Output C:\Mail\split -Mode Year

.EXAMPLE
    .\PstSplitter.ps1 -Source C:\Mail\archive.pst -Output C:\Mail\split -Mode Sender -EmailAddresses bob@example.com,amy@example.com -IncludeOther

.EXAMPLE
    .\PstSplitter.ps1 -Source C:\Mail\archive.pst -Output C:\Mail\split -Mode Month -WhatIf
    Shows how many items and MB would go to each file without writing anything.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$Source,
    [string]$Output,
    [ValidateSet('Size', 'Year', 'Month', 'Date', 'Folder', 'Sender')]
    [string]$Mode = 'Size',
    [int]$MaxSizeMB = 0,
    [datetime]$StartDate,
    [datetime]$EndDate,
    [string[]]$EmailAddresses,
    [ValidateSet('Sender', 'Recipients', 'Both')]
    [string]$MatchOn = 'Sender',
    [switch]$IncludeOther,
    [string[]]$ExcludeFolders = @('Deleted Items', 'Junk Email', 'Junk E-mail', 'Search Folders', 'Sync Issues', 'Conflicts', 'Local Failures', 'Server Failures'),
    [switch]$Gui
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# Outlook constants
$olStoreUnicode = 3
$olExchangeUserAddressEntry = 0
$olExchangeRemoteUserAddressEntry = 5
$olTo = 1
# Folder.DefaultItemType -> folder type to pass to Folders.Add
$FolderTypeForItemType = @{ 0 = 6; 1 = 9; 2 = 10; 3 = 13; 4 = 11; 5 = 12 }

function Release-Com($obj) {
    if ($null -ne $obj -and [System.Runtime.InteropServices.Marshal]::IsComObject($obj)) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj)
    }
}

function Get-SafeName([string]$name) {
    $bad = [System.IO.Path]::GetInvalidFileNameChars() -join ''
    $clean = ($name -replace "[$([regex]::Escape($bad))]", '_').Trim().TrimEnd('.')
    if ($clean.Length -gt 80) { $clean = $clean.Substring(0, 80) }
    if (-not $clean) { $clean = 'unnamed' }
    return $clean
}

function Get-ItemDate($item) {
    foreach ($prop in 'ReceivedTime', 'SentOn', 'Start', 'CreationTime') {
        try {
            $d = $item.$prop
            # Outlook uses 1/1/4501 for "none"
            if ($d -is [datetime] -and $d.Year -lt 4500) { return $d }
        } catch { }
    }
    return $null
}

function Get-SmtpAddress($addressEntry, [string]$fallback) {
    try {
        if ($null -ne $addressEntry) {
            $t = $addressEntry.AddressEntryUserType
            if ($t -eq $olExchangeUserAddressEntry -or $t -eq $olExchangeRemoteUserAddressEntry) {
                $ex = $addressEntry.GetExchangeUser()
                if ($null -ne $ex -and $ex.PrimarySmtpAddress) { return $ex.PrimarySmtpAddress.ToLower() }
            }
            if ($addressEntry.Address -and $addressEntry.Address -like '*@*') { return $addressEntry.Address.ToLower() }
        }
    } catch { }
    if ($fallback) { return $fallback.ToLower() }
    return $null
}

function Get-ItemAddresses($item, [string]$matchOn) {
    $list = New-Object System.Collections.Generic.List[string]
    if ($matchOn -ne 'Recipients') {
        $sender = $null
        try { $sender = Get-SmtpAddress $item.Sender $item.SenderEmailAddress } catch {
            try { if ($item.SenderEmailAddress) { $sender = $item.SenderEmailAddress.ToLower() } } catch { }
        }
        if ($sender) { $list.Add($sender) }
    }
    if ($matchOn -ne 'Sender') {
        try {
            foreach ($r in $item.Recipients) {
                $a = Get-SmtpAddress $r.AddressEntry $r.Address
                if ($a) { $list.Add($a) }
                Release-Com $r
            }
        } catch { }
    }
    return ,$list
}

# Works out which group (output file name suffix) an item belongs to.
# Returns $null when the item should be skipped.
function Get-GroupKey($item, [string[]]$folderPath, $opts) {
    switch ($opts.Mode) {
        'Size' { return '' }
        'Folder' { return (Get-SafeName $folderPath[0]) }
        'Year' {
            $d = Get-ItemDate $item
            if ($null -eq $d) { return 'NoDate' }
            return $d.ToString('yyyy')
        }
        'Month' {
            $d = Get-ItemDate $item
            if ($null -eq $d) { return 'NoDate' }
            return $d.ToString('yyyy-MM')
        }
        'Date' {
            $d = Get-ItemDate $item
            $inside = $null -ne $d -and $d -ge $opts.StartDate -and $d -lt $opts.EndDate.Date.AddDays(1)
            if ($inside) { return ('{0:yyyy-MM-dd}_to_{1:yyyy-MM-dd}' -f $opts.StartDate, $opts.EndDate) }
            if ($opts.IncludeOther) { return 'Other' }
            return $null
        }
        'Sender' {
            $addrs = Get-ItemAddresses $item $opts.MatchOn
            if ($opts.EmailAddresses.Count -gt 0) {
                foreach ($a in $addrs) {
                    if ($opts.EmailAddresses -contains $a) { return (Get-SafeName $a) }
                }
                if ($opts.IncludeOther) { return 'Other' }
                return $null
            }
            if ($addrs.Count -gt 0) { return (Get-SafeName $addrs[0]) }
            return 'Unknown'
        }
    }
}

function Find-Store($ns, [string]$path) {
    foreach ($s in $ns.Stores) {
        try {
            if ($s.FilePath -and ([System.IO.Path]::GetFullPath($s.FilePath) -ieq $path)) { return $s }
        } catch { }
    }
    return $null
}

function Get-OrAddFolder($parent, [string]$name, [int]$itemType) {
    foreach ($f in $parent.Folders) {
        if ($f.Name -ieq $name) { return $f }
    }
    $type = $FolderTypeForItemType[$itemType]
    if ($null -eq $type) { $type = 6 }
    return $parent.Folders.Add($name, $type)
}

# Walks the source PST and returns every folder that holds items, with its
# path below the root, e.g. @('Inbox','Clients').
function Get-SourceFolders($folder, [string[]]$path, [string[]]$exclude, $result) {
    foreach ($sub in $folder.Folders) {
        $name = $sub.Name
        if ($path.Count -eq 0 -and $exclude -contains $name) { continue }
        $subPath = @($path) + $name
        $result.Add([pscustomobject]@{ Folder = $sub; Path = $subPath; ItemType = $sub.DefaultItemType })
        Get-SourceFolders $sub $subPath $exclude $result
    }
}

function Invoke-PstSplit {
    param($opts, [scriptblock]$Log, [switch]$DryRun, [scriptblock]$Progress, [scriptblock]$ShouldCancel)

    $source = [System.IO.Path]::GetFullPath($opts.Source)
    if (-not (Test-Path -LiteralPath $source)) { throw "Source PST not found: $source" }
    if (-not $DryRun) {
        if (-not (Test-Path -LiteralPath $opts.Output)) { New-Item -ItemType Directory -Path $opts.Output | Out-Null }
    }
    $outDir = [System.IO.Path]::GetFullPath($opts.Output)
    $maxBytes = [int64]$opts.MaxSizeMB * 1MB
    $baseName = Get-SafeName ([System.IO.Path]::GetFileNameWithoutExtension($source))

    & $Log "Starting Outlook..."
    $outlook = New-Object -ComObject Outlook.Application
    $ns = $outlook.GetNamespace('MAPI')

    $srcStore = Find-Store $ns $source
    $weOpenedSource = $false
    if ($null -eq $srcStore) {
        & $Log "Opening $source"
        $ns.AddStoreEx($source, $olStoreUnicode)
        $srcStore = Find-Store $ns $source
        $weOpenedSource = $true
    }
    if ($null -eq $srcStore) { throw "Outlook could not open $source" }
    $srcStoreId = $srcStore.StoreID

    $folders = New-Object System.Collections.Generic.List[object]
    Get-SourceFolders $srcStore.GetRootFolder() @() $opts.ExcludeFolders $folders
    $total = 0
    foreach ($f in $folders) { $total += $f.Folder.Items.Count }
    & $Log ("Found {0} folders and {1} items" -f $folders.Count, $total)

    # key -> state of the output file currently being filled for that group
    $groups = @{}
    $opened = New-Object System.Collections.Generic.List[object]
    $done = 0; $copied = 0; $skipped = 0; $failed = 0; $cancelled = $false

    # Returns the output file for a group, starting the next part when the
    # current one would go over the size limit.
    function Get-Target([string]$key, [int64]$size) {
        $g = $groups[$key]
        if ($null -ne $g -and $maxBytes -gt 0 -and $g.Bytes -gt 0 -and ($g.Bytes + $size) -gt $maxBytes) {
            $g = New-Group $key ($g.Part + 1)
        }
        if ($null -eq $g) { $g = New-Group $key 1 }
        return $g
    }

    function New-Group([string]$key, [int]$part) {
        $name = $baseName
        if ($key) { $name += "_$key" }
        if ($maxBytes -gt 0) { $name += ('_part{0:000}' -f $part) }
        $path = Join-Path $outDir "$name.pst"
        $g = [pscustomobject]@{ Key = $key; Part = $part; Path = $path; Bytes = [int64]0; Items = 0; Store = $null; Root = $null; Folders = @{} }
        if (-not $DryRun) {
            if (Test-Path -LiteralPath $path) { throw "Output file already exists: $path. Choose an empty output folder." }
            & $Log "Creating $path"
            $ns.AddStoreEx($path, $olStoreUnicode)
            $g.Store = Find-Store $ns $path
            if ($null -eq $g.Store) { throw "Outlook could not create $path" }
            $g.Root = $g.Store.GetRootFolder()
            try { $g.Root.Name = $name } catch { }
        }
        $groups[$key] = $g
        $opened.Add($g)
        return $g
    }

    try {
        :folderLoop foreach ($f in $folders) {
            $pathText = $f.Path -join '\'
            # Collect IDs first: copying adds items to the folder while we work.
            $ids = New-Object System.Collections.Generic.List[string]
            $items = $f.Folder.Items
            foreach ($it in $items) {
                try { $ids.Add($it.EntryID) } catch { }
                Release-Com $it
            }
            Release-Com $items
            if ($ids.Count -eq 0) { continue }
            & $Log ("{0}: {1} items" -f $pathText, $ids.Count)

            foreach ($id in $ids) {
                if ($null -ne $ShouldCancel -and (& $ShouldCancel)) { $cancelled = $true; break folderLoop }
                $done++
                $item = $null; $copy = $null; $moved = $null
                try {
                    $item = $ns.GetItemFromID($id, $srcStoreId)
                    $key = Get-GroupKey $item $f.Path $opts
                    if ($null -eq $key) { $skipped++; continue }
                    $size = [int64]0
                    try { $size = [int64]$item.Size } catch { }
                    $g = Get-Target $key $size

                    if (-not $DryRun) {
                        $dest = $g.Folders[$pathText]
                        if ($null -eq $dest) {
                            $dest = $g.Root
                            for ($i = 0; $i -lt $f.Path.Count; $i++) {
                                $dest = Get-OrAddFolder $dest $f.Path[$i] $f.ItemType
                            }
                            $g.Folders[$pathText] = $dest
                        }
                        $copy = $item.Copy()
                        $moved = $copy.Move($dest)
                    }
                    $g.Bytes += $size
                    $g.Items++
                    $copied++
                } catch {
                    $failed++
                    # Don't leave a stray copy in the source if the move failed.
                    if ($null -ne $copy -and $null -eq $moved) { try { $copy.Delete() } catch { } }
                    & $Log ("  Could not copy an item in {0}: {1}" -f $pathText, $_.Exception.Message)
                } finally {
                    Release-Com $moved; Release-Com $copy; Release-Com $item
                }
                if ($null -ne $Progress) { & $Progress $done $total $pathText }
                if ($done % 200 -eq 0) {
                    if ($null -eq $Progress) { & $Log ("  {0} of {1} items processed" -f $done, $total) }
                    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
                }
            }
        }
    } finally {
        foreach ($g in $opened) {
            if ($null -ne $g.Root) {
                try { $ns.RemoveStore($g.Root) } catch { }
                foreach ($k in @($g.Folders.Keys)) { Release-Com $g.Folders[$k] }
                Release-Com $g.Root; Release-Com $g.Store
            }
        }
        if ($weOpenedSource) {
            try { $ns.RemoveStore($srcStore.GetRootFolder()) } catch { }
        }
        foreach ($f in $folders) { Release-Com $f.Folder }
        Release-Com $srcStore; Release-Com $ns; Release-Com $outlook
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }

    & $Log ''
    if ($cancelled) { & $Log 'Stopped. The files written so far are complete and can be opened.' }
    if ($DryRun) { & $Log 'Preview only. No files were written.' }
    $files = New-Object System.Collections.Generic.List[object]
    foreach ($g in $opened) {
        $name = [System.IO.Path]::GetFileName($g.Path)
        $files.Add([pscustomobject]@{ Name = $name; Path = $g.Path; Items = $g.Items; MB = $g.Bytes / 1MB })
        & $Log ("{0}  {1} items, {2:N1} MB" -f $name, $g.Items, ($g.Bytes / 1MB))
    }
    & $Log ("Done. Copied {0}, skipped {1}, failed {2}." -f $copied, $skipped, $failed)
    return [pscustomobject]@{
        Files = $files; Copied = $copied; Skipped = $skipped; Failed = $failed
        Cancelled = $cancelled; Output = $outDir; Preview = [bool]$DryRun
    }
}

function Test-Options($opts) {
    if (-not $opts.Source) { return 'Choose the PST file to split.' }
    if (-not $opts.Output) { return 'Choose an output folder.' }
    if ($opts.Mode -eq 'Size' -and $opts.MaxSizeMB -le 0) { return 'Enter a maximum size in MB.' }
    if ($opts.Mode -eq 'Date') {
        if ($null -eq $opts.StartDate -or $null -eq $opts.EndDate) { return 'Enter a start and end date.' }
        if ($opts.EndDate -lt $opts.StartDate) { return 'The end date is before the start date.' }
    }
    return $null
}

function New-Options {
    param($Source, $Output, $Mode, $MaxSizeMB, $StartDate, $EndDate, $EmailAddresses, $MatchOn, $IncludeOther, $ExcludeFolders)
    $emails = @()
    if ($EmailAddresses) {
        $emails = @($EmailAddresses -split '[,;\s]+' | Where-Object { $_ } | ForEach-Object { $_.Trim().ToLower() })
    }
    return [pscustomobject]@{
        Source = $Source; Output = $Output; Mode = $Mode; MaxSizeMB = [int]$MaxSizeMB
        StartDate = $StartDate; EndDate = $EndDate; EmailAddresses = $emails; MatchOn = $MatchOn
        IncludeOther = [bool]$IncludeOther; ExcludeFolders = @($ExcludeFolders)
    }
}

function Show-SplitterWindow {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    function P([int]$x, [int]$y) { New-Object System.Drawing.Point($x, $y) }
    function Sz([int]$w, [int]$h) { New-Object System.Drawing.Size($w, $h) }
    function Add-Ctl($parent, [string]$type, [int]$x, [int]$y, [int]$w, [int]$h, [string]$text) {
        $c = New-Object "System.Windows.Forms.$type"
        $c.Location = P $x $y
        if ($w -gt 0) { $c.Width = $w }
        if ($h -gt 0) { $c.Height = $h }
        if ($text) { $c.Text = $text }
        $parent.Controls.Add($c)
        return $c
    }
    function Add-Label($parent, [string]$text, [int]$x, [int]$y, [switch]$Hint) {
        $l = Add-Ctl $parent 'Label' $x ($y + 3) 0 0 $text
        $l.AutoSize = $true
        if ($Hint) { $l.ForeColor = [System.Drawing.Color]::DimGray }
        return $l
    }

    $state = @{ Running = $false; Cancel = $false; CloseAfter = $false; Output = $null }
    $font = New-Object System.Drawing.Font('Segoe UI', 9)
    $accent = [System.Drawing.Color]::FromArgb(0, 84, 147)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'PST Splitter'
    $form.ClientSize = Sz 720 740
    $form.MinimumSize = Sz 736 700
    $form.StartPosition = 'CenterScreen'
    $form.Font = $font
    $form.AllowDrop = $true
    try { $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon((Get-Command powershell.exe).Source) } catch { }

    # Header
    $header = New-Object System.Windows.Forms.Panel
    $header.Dock = 'Top'; $header.Height = 64; $header.BackColor = $accent
    $form.Controls.Add($header)
    $title = Add-Label $header 'PST Splitter' 16 6
    $title.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 16)
    $title.ForeColor = [System.Drawing.Color]::White
    $sub = Add-Label $header 'Split a large Outlook data file into smaller ones. The original is not changed.' 18 38
    $sub.ForeColor = [System.Drawing.Color]::FromArgb(210, 228, 245)

    # 1. Files
    $gbFiles = Add-Ctl $form 'GroupBox' 12 74 696 104 '1. Files'
    $gbFiles.Anchor = 'Top, Left, Right'
    Add-Label $gbFiles 'PST file to split' 12 24 | Out-Null
    $txtSource = Add-Ctl $gbFiles 'TextBox' 130 24 456 0 ''
    $txtSource.Anchor = 'Top, Left, Right'
    $btnSource = Add-Ctl $gbFiles 'Button' 594 23 90 26 'Browse...'
    $btnSource.Anchor = 'Top, Right'
    $lblSourceInfo = Add-Label $gbFiles 'Or drag a .pst file onto this window.' 130 46 -Hint
    Add-Label $gbFiles 'Save split files to' 12 70 | Out-Null
    $txtOutput = Add-Ctl $gbFiles 'TextBox' 130 70 456 0 ''
    $txtOutput.Anchor = 'Top, Left, Right'
    $btnOutput = Add-Ctl $gbFiles 'Button' 594 69 90 26 'Browse...'
    $btnOutput.Anchor = 'Top, Right'

    # 2. Split by
    $gbSplit = Add-Ctl $form 'GroupBox' 12 186 696 240 '2. Split by'
    $gbSplit.Anchor = 'Top, Left, Right'
    $modeList = @(
        @{ Key = 'Size'; Text = 'Size'; Help = 'Starts a new file each time the current one reaches the size below. Files are named _part001, _part002, and so on.' },
        @{ Key = 'Year'; Text = 'Year'; Help = 'One file per year, using the date each item was received or sent.' },
        @{ Key = 'Month'; Text = 'Month'; Help = 'One file per month, using the date each item was received or sent.' },
        @{ Key = 'Date'; Text = 'Date range'; Help = 'One file with the items between these dates (both days included).' },
        @{ Key = 'Folder'; Text = 'Folder'; Help = 'One file per top-level folder: Inbox, Sent Items, Calendar, Contacts and so on. Subfolders go with their parent.' },
        @{ Key = 'Sender'; Text = 'Email address'; Help = 'One file per email address. List the addresses to pull out, or leave the box empty for one file per sender.' }
    )
    $radios = @()
    $panels = @{}
    $i = 0
    foreach ($m in $modeList) {
        $rb = Add-Ctl $gbSplit 'RadioButton' 16 (26 + 27 * $i) 150 24 $m.Text
        $rb.Tag = $m.Key
        $radios += $rb
        $pnl = Add-Ctl $gbSplit 'Panel' 180 20 504 150 ''
        $pnl.Anchor = 'Top, Left, Right'
        $pnl.Visible = $false
        $help = Add-Label $pnl $m.Help 0 0 -Hint
        $help.AutoSize = $false; $help.Size = Sz 500 40
        $panels[$m.Key] = $pnl
        $i++
    }
    $sep = Add-Ctl $gbSplit 'Label' 172 22 2 150 ''
    $sep.BorderStyle = 'Fixed3D'

    # Date range fields
    $pDate = $panels['Date']
    Add-Label $pDate 'From' 0 48 | Out-Null
    $dtStart = Add-Ctl $pDate 'DateTimePicker' 44 48 120 0 ''
    $dtStart.Format = 'Short'; $dtStart.Value = (Get-Date).AddYears(-1).Date
    Add-Label $pDate 'to' 174 48 | Out-Null
    $dtEnd = Add-Ctl $pDate 'DateTimePicker' 198 48 120 0 ''
    $dtEnd.Format = 'Short'
    $chkOtherDate = Add-Ctl $pDate 'CheckBox' 0 84 400 24 'Put items outside the dates into an "_Other" file'

    # Email address fields
    $pSender = $panels['Sender']
    Add-Label $pSender 'Addresses' 0 44 | Out-Null
    $txtEmail = Add-Ctl $pSender 'TextBox' 70 44 430 0 ''
    $txtEmail.Multiline = $true; $txtEmail.AcceptsReturn = $true; $txtEmail.ScrollBars = 'Vertical'; $txtEmail.Height = 46
    $txtEmail.Anchor = 'Top, Left, Right'
    Add-Label $pSender 'One per line, or separated by commas.' 70 90 -Hint | Out-Null
    Add-Label $pSender 'Match on' 0 116 | Out-Null
    $cboMatch = Add-Ctl $pSender 'ComboBox' 70 116 150 0 ''
    $cboMatch.DropDownStyle = 'DropDownList'
    [void]$cboMatch.Items.AddRange(@('Sender', 'Recipients', 'Both'))
    $cboMatch.SelectedIndex = 0
    $chkOtherSender = Add-Ctl $pSender 'CheckBox' 236 116 260 24 'Unmatched items go to an "_Other" file'

    # Shared options
    $line = Add-Ctl $gbSplit 'Label' 12 180 672 2 ''
    $line.BorderStyle = 'Fixed3D'; $line.Anchor = 'Top, Left, Right'
    $chkSize = Add-Ctl $gbSplit 'CheckBox' 16 190 190 24 'Limit each file to'
    $numSize = Add-Ctl $gbSplit 'NumericUpDown' 206 190 90 0 ''
    $numSize.Minimum = 10; $numSize.Maximum = 50000; $numSize.Increment = 256; $numSize.Value = 2048
    $numSize.ThousandsSeparator = $true
    Add-Label $gbSplit 'MB   (set about 10% under the size you need)' 300 190 -Hint | Out-Null
    $chkSkip = Add-Ctl $gbSplit 'CheckBox' 16 212 460 24 'Skip Deleted Items, Junk Email and Outlook sync folders'
    $chkSkip.Checked = $true

    # 3. Run
    $btnPreview = Add-Ctl $form 'Button' 12 438 110 32 'Preview'
    $btnSplit = Add-Ctl $form 'Button' 130 438 110 32 'Split'
    $btnSplit.BackColor = $accent; $btnSplit.ForeColor = [System.Drawing.Color]::White
    $btnSplit.FlatStyle = 'Flat'; $btnSplit.FlatAppearance.BorderSize = 0
    $btnSplit.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
    $btnStop = Add-Ctl $form 'Button' 248 438 90 32 'Stop'
    $btnStop.Enabled = $false
    $btnOpen = Add-Ctl $form 'Button' 568 438 140 32 'Open output folder'
    $btnOpen.Anchor = 'Top, Right'; $btnOpen.Enabled = $false
    $tip = New-Object System.Windows.Forms.ToolTip
    $tip.SetToolTip($btnPreview, 'Shows the files that would be made, without writing anything')
    $tip.SetToolTip($btnSplit, 'Writes the new PST files')

    $bar = Add-Ctl $form 'ProgressBar' 12 480 696 18 ''
    $bar.Anchor = 'Top, Left, Right'
    $lblStatus = Add-Label $form 'Choose a PST file to start.' 12 500
    $lblStatus.AutoSize = $false; $lblStatus.Size = Sz 696 20
    $lblStatus.Anchor = 'Top, Left, Right'; $lblStatus.AutoEllipsis = $true

    $tabs = Add-Ctl $form 'TabControl' 12 526 696 202 ''
    $tabs.Anchor = 'Top, Bottom, Left, Right'
    $tabResults = New-Object System.Windows.Forms.TabPage 'Results'
    $tabLog = New-Object System.Windows.Forms.TabPage 'Log'
    $tabs.TabPages.AddRange(@($tabResults, $tabLog))

    $list = New-Object System.Windows.Forms.ListView
    $list.Dock = 'Fill'; $list.View = 'Details'; $list.FullRowSelect = $true; $list.GridLines = $true
    [void]$list.Columns.Add('File', 420)
    [void]$list.Columns.Add('Items', 100, 'Right')
    [void]$list.Columns.Add('Size (MB)', 120, 'Right')
    $tabResults.Controls.Add($list)

    $txtLog = New-Object System.Windows.Forms.TextBox
    $txtLog.Dock = 'Fill'; $txtLog.Multiline = $true; $txtLog.ScrollBars = 'Vertical'; $txtLog.ReadOnly = $true
    $txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)
    $tabLog.Controls.Add($txtLog)

    # --- behaviour ---------------------------------------------------------

    function Get-SelectedMode {
        foreach ($rb in $radios) { if ($rb.Checked) { return [string]$rb.Tag } }
        return 'Size'
    }

    function Update-ModeView {
        $m = Get-SelectedMode
        foreach ($k in $panels.Keys) { $panels[$k].Visible = ($k -eq $m) }
        if ($m -eq 'Size') { $chkSize.Checked = $true; $chkSize.Enabled = $false } else { $chkSize.Enabled = -not $state.Running }
        $numSize.Enabled = $chkSize.Checked
    }

    function Set-Source([string]$path) {
        $txtSource.Text = $path
        if (-not $txtOutput.Text -and $path) {
            $txtOutput.Text = Join-Path ([System.IO.Path]::GetDirectoryName($path)) (([System.IO.Path]::GetFileNameWithoutExtension($path)) + '_split')
        }
        Show-SourceInfo
    }

    function Show-SourceInfo {
        $path = $txtSource.Text.Trim()
        $exists = $false
        try { $exists = $path -and (Test-Path -LiteralPath $path -PathType Leaf) } catch { }
        if ($exists) {
            $mb = (Get-Item -LiteralPath $path).Length / 1MB
            $lblSourceInfo.Text = '{0:N0} MB. Click Preview to see how it would be split.' -f $mb
            $lblStatus.Text = 'Ready.'
        } elseif ($path) {
            $lblSourceInfo.Text = 'File not found.'
        } else {
            $lblSourceInfo.Text = 'Or drag a .pst file onto this window.'
        }
    }

    function Set-Running([bool]$running) {
        $state.Running = $running
        foreach ($c in @($gbFiles, $gbSplit, $btnPreview, $btnSplit)) { $c.Enabled = -not $running }
        $btnStop.Enabled = $running
        if ($running) { $btnOpen.Enabled = $false }
        $form.UseWaitCursor = $running
        $btnStop.UseWaitCursor = $false
    }

    function Write-Log([string]$msg) {
        $txtLog.AppendText("$msg`r`n")
        if ($msg -and $msg -notlike '  *' -and $msg -notlike '*items, *MB') { $lblStatus.Text = $msg }
        [System.Windows.Forms.Application]::DoEvents()
    }

    function Start-Split([bool]$preview) {
        $size = 0
        if ($chkSize.Checked) { $size = [int]$numSize.Value }
        $exclude = @()
        if ($chkSkip.Checked) { $exclude = $ExcludeFolders }
        $mode = Get-SelectedMode
        $other = ($mode -eq 'Date' -and $chkOtherDate.Checked) -or ($mode -eq 'Sender' -and $chkOtherSender.Checked)
        $opts = New-Options -Source $txtSource.Text.Trim() -Output $txtOutput.Text.Trim() `
            -Mode $mode -MaxSizeMB $size -StartDate $dtStart.Value.Date -EndDate $dtEnd.Value.Date `
            -EmailAddresses $txtEmail.Text -MatchOn ([string]$cboMatch.SelectedItem) `
            -IncludeOther $other -ExcludeFolders $exclude
        $problem = Test-Options $opts
        if (-not $problem -and -not (Test-Path -LiteralPath $opts.Source -PathType Leaf)) { $problem = 'The PST file was not found.' }
        if ($problem) {
            [void][System.Windows.Forms.MessageBox]::Show($form, $problem, 'PST Splitter', 'OK', 'Warning')
            return
        }
        if (-not $preview) {
            if ((Test-Path -LiteralPath $opts.Output) -and (Get-ChildItem -LiteralPath $opts.Output -Filter *.pst | Select-Object -First 1)) {
                [void][System.Windows.Forms.MessageBox]::Show($form, "The output folder already has PST files in it.`n`nChoose an empty folder.", 'PST Splitter', 'OK', 'Warning')
                return
            }
            $msg = "Split $([System.IO.Path]::GetFileName($opts.Source)) into $($opts.Output)?`n`n" +
                "Outlook will open in the background while it works. Keep a backup of the original PST until you've checked the new files."
            if ([System.Windows.Forms.MessageBox]::Show($form, $msg, 'PST Splitter', 'OKCancel', 'Question') -ne 'OK') { return }
        }

        $txtLog.Clear(); $list.Items.Clear()
        $bar.Value = 0
        $state.Cancel = $false
        Set-Running $true
        $onProgress = {
            param($done, $total, $folder)
            if ($total -gt 0) {
                $bar.Maximum = $total
                $bar.Value = [Math]::Min($done, $total)
                $lblStatus.Text = '{0} {1:N0} of {2:N0} items ({3:P0})  -  {4}' -f $(if ($preview) { 'Checking' } else { 'Copying' }), $done, $total, ($done / $total), $folder
            }
            [System.Windows.Forms.Application]::DoEvents()
        }
        try {
            $result = Invoke-PstSplit -opts $opts -Log { param($m) Write-Log $m } -DryRun:$preview `
                -Progress $onProgress -ShouldCancel { $state.Cancel }
            foreach ($f in $result.Files) {
                $row = New-Object System.Windows.Forms.ListViewItem($f.Name)
                [void]$row.SubItems.Add(('{0:N0}' -f $f.Items))
                [void]$row.SubItems.Add(('{0:N1}' -f $f.MB))
                $row.Tag = $f.Path
                [void]$list.Items.Add($row)
            }
            $tabs.SelectedTab = $tabResults
            $summary = '{0:N0} items in {1} file(s)' -f $result.Copied, $result.Files.Count
            if ($result.Skipped) { $summary += (', {0:N0} skipped' -f $result.Skipped) }
            if ($result.Failed) { $summary += (', {0:N0} could not be copied (see Log)' -f $result.Failed) }
            if ($result.Cancelled) { $lblStatus.Text = "Stopped. $summary." }
            elseif ($preview) { $lblStatus.Text = "Preview: $summary. Nothing was written." }
            else {
                $bar.Value = $bar.Maximum
                $lblStatus.Text = "Finished: $summary."
            }
            if (-not $preview -and $result.Files.Count -gt 0) {
                $state.Output = $result.Output
                $btnOpen.Enabled = $true
            }
        } catch {
            Write-Log "ERROR: $($_.Exception.Message)"
            $lblStatus.Text = 'Something went wrong. See the Log tab.'
            $tabs.SelectedTab = $tabLog
            [void][System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'PST Splitter', 'OK', 'Error')
        } finally {
            Set-Running $false
            Update-ModeView
        }
        if ($state.CloseAfter) { $form.Close() }
    }

    foreach ($rb in $radios) { $rb.Add_CheckedChanged({ Update-ModeView }) }
    $radios[0].Checked = $true
    $chkSize.Add_CheckedChanged({ $numSize.Enabled = $chkSize.Checked })

    $btnSource.Add_Click({
        $d = New-Object System.Windows.Forms.OpenFileDialog
        $d.Filter = 'Outlook data files (*.pst)|*.pst|All files (*.*)|*.*'
        if ($d.ShowDialog($form) -eq 'OK') { Set-Source $d.FileName }
    })
    $btnOutput.Add_Click({
        $d = New-Object System.Windows.Forms.FolderBrowserDialog
        $d.Description = 'Choose where to save the split PST files'
        if ($txtOutput.Text -and (Test-Path -LiteralPath $txtOutput.Text)) { $d.SelectedPath = $txtOutput.Text }
        if ($d.ShowDialog($form) -eq 'OK') { $txtOutput.Text = $d.SelectedPath }
    })
    $txtSource.Add_Leave({ Show-SourceInfo })

    $form.Add_DragEnter({
        param($s, $e)
        if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop) -and -not $state.Running) { $e.Effect = 'Copy' }
    })
    $form.Add_DragDrop({
        param($s, $e)
        $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        if ($files -and $files.Count -gt 0) { Set-Source ([string]$files[0]) }
    })

    $btnPreview.Add_Click({ Start-Split $true })
    $btnSplit.Add_Click({ Start-Split $false })
    $btnStop.Add_Click({
        $state.Cancel = $true
        $btnStop.Enabled = $false
        $lblStatus.Text = 'Stopping after the current item...'
    })
    $btnOpen.Add_Click({ if ($state.Output) { Start-Process explorer.exe $state.Output } })
    $list.Add_DoubleClick({
        if ($list.SelectedItems.Count -gt 0 -and $list.SelectedItems[0].Tag -and (Test-Path -LiteralPath $list.SelectedItems[0].Tag)) {
            Start-Process explorer.exe ('/select,"{0}"' -f $list.SelectedItems[0].Tag)
        }
    })

    $form.Add_FormClosing({
        param($s, $e)
        if ($state.Running) {
            $answer = [System.Windows.Forms.MessageBox]::Show($form, 'A split is running. Stop it and close?', 'PST Splitter', 'YesNo', 'Question')
            if ($answer -eq 'Yes') { $state.Cancel = $true; $state.CloseAfter = $true }
            $e.Cancel = $true
        }
    })

    $form.AcceptButton = $btnPreview
    [void]$form.ShowDialog()
}

# ---------------------------------------------------------------------------

if ($Gui -or -not $Source) {
    Show-SplitterWindow
    return
}

if (-not $Output) {
    $Output = Join-Path ([System.IO.Path]::GetDirectoryName((Resolve-Path -LiteralPath $Source).Path)) (([System.IO.Path]::GetFileNameWithoutExtension($Source)) + '_split')
}
$sd = $null; $ed = $null
if ($PSBoundParameters.ContainsKey('StartDate')) { $sd = $StartDate }
if ($PSBoundParameters.ContainsKey('EndDate')) { $ed = $EndDate }
$options = New-Options -Source $Source -Output $Output -Mode $Mode -MaxSizeMB $MaxSizeMB `
    -StartDate $sd -EndDate $ed -EmailAddresses ($EmailAddresses -join ',') `
    -MatchOn $MatchOn -IncludeOther $IncludeOther.IsPresent -ExcludeFolders $ExcludeFolders
$problem = Test-Options $options
if ($problem) { throw $problem }

$dry = $WhatIfPreference -eq $true
$null = Invoke-PstSplit -opts $options -Log { param($m) Write-Host $m } -DryRun:$dry
