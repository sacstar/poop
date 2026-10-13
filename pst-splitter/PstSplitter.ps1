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
    param($opts, [scriptblock]$Log, [switch]$DryRun)

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
    $done = 0; $copied = 0; $skipped = 0; $failed = 0

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
        foreach ($f in $folders) {
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
                if ($done % 200 -eq 0) {
                    & $Log ("  {0} of {1} items processed" -f $done, $total)
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
    if ($DryRun) { & $Log 'Preview only. No files were written.' }
    foreach ($g in $opened) {
        & $Log ("{0}  {1} items, {2:N1} MB" -f [System.IO.Path]::GetFileName($g.Path), $g.Items, ($g.Bytes / 1MB))
    }
    & $Log ("Done. Copied {0}, skipped {1}, failed {2}." -f $copied, $skipped, $failed)
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

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'PST Splitter'
    $form.Size = New-Object System.Drawing.Size(640, 600)
    $form.StartPosition = 'CenterScreen'
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $y = 15
    function Add-Label($text, $x, $yy) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $text; $l.Location = New-Object System.Drawing.Point($x, ($yy + 3)); $l.AutoSize = $true
        $form.Controls.Add($l); return $l
    }

    Add-Label 'PST file to split' 15 $y | Out-Null
    $txtSource = New-Object System.Windows.Forms.TextBox
    $txtSource.Location = New-Object System.Drawing.Point(140, $y); $txtSource.Width = 380
    $form.Controls.Add($txtSource)
    $btnSource = New-Object System.Windows.Forms.Button
    $btnSource.Text = 'Browse...'; $btnSource.Location = New-Object System.Drawing.Point(530, ($y - 1)); $btnSource.Width = 80
    $form.Controls.Add($btnSource)
    $btnSource.Add_Click({
        $d = New-Object System.Windows.Forms.OpenFileDialog
        $d.Filter = 'Outlook data files (*.pst)|*.pst'
        if ($d.ShowDialog() -eq 'OK') {
            $txtSource.Text = $d.FileName
            if (-not $txtOutput.Text) {
                $txtOutput.Text = Join-Path ([System.IO.Path]::GetDirectoryName($d.FileName)) (([System.IO.Path]::GetFileNameWithoutExtension($d.FileName)) + '_split')
            }
        }
    })

    $y += 35
    Add-Label 'Save split files to' 15 $y | Out-Null
    $txtOutput = New-Object System.Windows.Forms.TextBox
    $txtOutput.Location = New-Object System.Drawing.Point(140, $y); $txtOutput.Width = 380
    $form.Controls.Add($txtOutput)
    $btnOutput = New-Object System.Windows.Forms.Button
    $btnOutput.Text = 'Browse...'; $btnOutput.Location = New-Object System.Drawing.Point(530, ($y - 1)); $btnOutput.Width = 80
    $form.Controls.Add($btnOutput)
    $btnOutput.Add_Click({
        $d = New-Object System.Windows.Forms.FolderBrowserDialog
        if ($d.ShowDialog() -eq 'OK') { $txtOutput.Text = $d.SelectedPath }
    })

    $y += 35
    Add-Label 'Split by' 15 $y | Out-Null
    $cboMode = New-Object System.Windows.Forms.ComboBox
    $cboMode.DropDownStyle = 'DropDownList'
    $cboMode.Location = New-Object System.Drawing.Point(140, $y); $cboMode.Width = 250
    $modes = [ordered]@{
        'Size' = 'Size'; 'Year' = 'Year'; 'Month' = 'Month'; 'Date range' = 'Date'
        'Top-level folder' = 'Folder'; 'Email address' = 'Sender'
    }
    foreach ($k in $modes.Keys) { [void]$cboMode.Items.Add($k) }
    $cboMode.SelectedIndex = 0
    $form.Controls.Add($cboMode)

    $y += 35
    $chkSize = New-Object System.Windows.Forms.CheckBox
    $chkSize.Text = 'Max size per file (MB)'; $chkSize.AutoSize = $true
    $chkSize.Location = New-Object System.Drawing.Point(140, $y); $chkSize.Checked = $true
    $form.Controls.Add($chkSize)
    $numSize = New-Object System.Windows.Forms.NumericUpDown
    $numSize.Location = New-Object System.Drawing.Point(320, $y); $numSize.Width = 90
    $numSize.Minimum = 10; $numSize.Maximum = 50000; $numSize.Value = 2048; $numSize.Increment = 256
    $form.Controls.Add($numSize)

    $y += 35
    $lblDates = Add-Label 'From' 140 $y
    $dtStart = New-Object System.Windows.Forms.DateTimePicker
    $dtStart.Format = 'Short'; $dtStart.Location = New-Object System.Drawing.Point(180, $y); $dtStart.Width = 110
    $dtStart.Value = (Get-Date).AddYears(-1).Date
    $form.Controls.Add($dtStart)
    $lblTo = Add-Label 'to' 300 $y
    $dtEnd = New-Object System.Windows.Forms.DateTimePicker
    $dtEnd.Format = 'Short'; $dtEnd.Location = New-Object System.Drawing.Point(325, $y); $dtEnd.Width = 110
    $form.Controls.Add($dtEnd)

    $y += 35
    $lblEmail = Add-Label 'Addresses' 140 $y
    $txtEmail = New-Object System.Windows.Forms.TextBox
    $txtEmail.Location = New-Object System.Drawing.Point(210, $y); $txtEmail.Width = 400
    $form.Controls.Add($txtEmail)
    $y += 28
    $lblEmailHint = Add-Label 'Separate with commas. Leave empty for one file per sender.' 210 $y
    $lblEmailHint.ForeColor = [System.Drawing.Color]::DimGray
    $y += 25
    $lblMatch = Add-Label 'Match' 140 $y
    $cboMatch = New-Object System.Windows.Forms.ComboBox
    $cboMatch.DropDownStyle = 'DropDownList'
    $cboMatch.Location = New-Object System.Drawing.Point(210, $y); $cboMatch.Width = 180
    [void]$cboMatch.Items.AddRange(@('Sender', 'Recipients', 'Both'))
    $cboMatch.SelectedIndex = 0
    $form.Controls.Add($cboMatch)

    $y += 35
    $chkOther = New-Object System.Windows.Forms.CheckBox
    $chkOther.Text = 'Put items that do not match into an "Other" file'; $chkOther.AutoSize = $true
    $chkOther.Location = New-Object System.Drawing.Point(140, $y)
    $form.Controls.Add($chkOther)

    $y += 28
    $chkSkip = New-Object System.Windows.Forms.CheckBox
    $chkSkip.Text = 'Skip Deleted Items, Junk Email and sync folders'; $chkSkip.AutoSize = $true
    $chkSkip.Location = New-Object System.Drawing.Point(140, $y); $chkSkip.Checked = $true
    $form.Controls.Add($chkSkip)

    $y += 35
    $btnPreview = New-Object System.Windows.Forms.Button
    $btnPreview.Text = 'Preview'; $btnPreview.Location = New-Object System.Drawing.Point(140, $y); $btnPreview.Width = 100
    $form.Controls.Add($btnPreview)
    $btnStart = New-Object System.Windows.Forms.Button
    $btnStart.Text = 'Split'; $btnStart.Location = New-Object System.Drawing.Point(250, $y); $btnStart.Width = 100
    $form.Controls.Add($btnStart)

    $y += 38
    $txtLog = New-Object System.Windows.Forms.TextBox
    $txtLog.Multiline = $true; $txtLog.ScrollBars = 'Vertical'; $txtLog.ReadOnly = $true
    $txtLog.Location = New-Object System.Drawing.Point(15, $y)
    $txtLog.Size = New-Object System.Drawing.Size(595, (545 - $y))
    $txtLog.Anchor = 'Top, Bottom, Left, Right'
    $txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtLog)

    $updateVisibility = {
        $m = $modes[[string]$cboMode.SelectedItem]
        $isDate = $m -eq 'Date'; $isSender = $m -eq 'Sender'
        foreach ($c in @($lblDates, $dtStart, $lblTo, $dtEnd)) { $c.Enabled = $isDate }
        foreach ($c in @($lblEmail, $txtEmail, $lblEmailHint, $lblMatch, $cboMatch)) { $c.Enabled = $isSender }
        $chkOther.Enabled = $isDate -or $isSender
        if ($m -eq 'Size') { $chkSize.Checked = $true; $chkSize.Enabled = $false } else { $chkSize.Enabled = $true }
        $numSize.Enabled = $chkSize.Checked
    }
    $cboMode.Add_SelectedIndexChanged($updateVisibility)
    $chkSize.Add_CheckedChanged({ $numSize.Enabled = $chkSize.Checked })
    & $updateVisibility

    $run = {
        param([bool]$preview)
        $size = 0
        if ($chkSize.Checked) { $size = [int]$numSize.Value }
        $exclude = @()
        if ($chkSkip.Checked) { $exclude = $ExcludeFolders }
        $opts = New-Options -Source $txtSource.Text.Trim() -Output $txtOutput.Text.Trim() `
            -Mode $modes[[string]$cboMode.SelectedItem] -MaxSizeMB $size `
            -StartDate $dtStart.Value.Date -EndDate $dtEnd.Value.Date -EmailAddresses $txtEmail.Text `
            -MatchOn ([string]$cboMatch.SelectedItem) -IncludeOther $chkOther.Checked -ExcludeFolders $exclude
        $problem = Test-Options $opts
        if ($problem) { [void][System.Windows.Forms.MessageBox]::Show($problem, 'PST Splitter'); return }

        $txtLog.Clear()
        $btnStart.Enabled = $false; $btnPreview.Enabled = $false
        $form.Cursor = 'WaitCursor'
        $log = {
            param($msg)
            $txtLog.AppendText("$msg`r`n")
            [System.Windows.Forms.Application]::DoEvents()
        }
        try {
            Invoke-PstSplit -opts $opts -Log $log -DryRun:$preview
        } catch {
            & $log "ERROR: $($_.Exception.Message)"
            [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'PST Splitter', 'OK', 'Error')
        } finally {
            $btnStart.Enabled = $true; $btnPreview.Enabled = $true
            $form.Cursor = 'Default'
        }
    }
    $btnPreview.Add_Click({ & $run $true })
    $btnStart.Add_Click({ & $run $false })

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
Invoke-PstSplit -opts $options -Log { param($m) Write-Host $m } -DryRun:$dry
