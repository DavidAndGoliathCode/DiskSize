# PowerShell script: list drives with size, used, free and percentages
# This script has been created in partnership with my AI Buddy
# this is a change test
# This is collaberator adding a test to show GIT
function Convert-Bytes {
    param([double]$Bytes)
    if ($Bytes -lt 1KB) { return "$Bytes B" }
    $units = "B","KB","MB","GB","TB","PB"
    $i = 0
    while ($Bytes -ge 1024 -and $i -lt $units.Length-1) {
        $Bytes /= 1024
        $i++
    }
    return "{0:N2} {1}" -f $Bytes, $units[$i]
}

$banner = @'
    _____  _             
 ____      _         
|  _ \ ___(_)_______ 
| | | / __| |_  / _ \
| |_| \__ \ |/ /  __/
|____/|___/_/___\___|

 Dsize - Disk sizes at a glance

 "Know your bytes before they bite!"
'@

Write-Host $banner -ForegroundColor Cyan

Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=2 OR DriveType=3" |
Select-Object DeviceID, VolumeName, @{Name='SizeBytes';Expression={[int64]$_.Size}}, @{Name='FreeBytes';Expression={[int64]$_.FreeSpace}} |
ForEach-Object {
    $size = $_.SizeBytes
    $free = $_.FreeBytes
    $used = $size - $free
    if ($size -gt 0) {
        $pctUsed = ($used / $size) * 100
        $pctFree = ($free / $size) * 100
    } else {
        $pctUsed = 0
        $pctFree = 0
    }

    # Determine encryption status (BitLocker) with fallback to manage-bde parsing
    $encLabel = 'Not Encrypted'
    $encPct = $null
    try {
        $bl = Get-BitLockerVolume -MountPoint $_.DeviceID -ErrorAction SilentlyContinue
        if ($bl) {
            # ProtectionStatus may be 0/1 or strings depending on environment; check common properties
            if ($bl.ProtectionStatus -ne $null) {
                if ($bl.ProtectionStatus -eq 'On' -or $bl.ProtectionStatus -eq 1) { $encLabel = 'Encrypted' }
                else { $encLabel = 'Not Encrypted' }
            }
            if ($encLabel -eq 'Encrypted' -and $bl.EncryptionPercentage -ne $null) {
                $encPct = [int]$bl.EncryptionPercentage
            } elseif ($bl.VolumeStatus -match 'FullyEncrypted') {
                $encPct = 100
                $encLabel = 'Encrypted'
            }
        } else {
            # fallback: try manage-bde output parsing
            $mb = & manage-bde -status $_.DeviceID 2>$null
            if ($mb) {
                $line = $mb | Select-String -Pattern 'Percentage Encrypted' -SimpleMatch
                if ($line) {
                    $val = ($line -split ':')[-1].Trim().TrimEnd('%')
                    $parsed = 0
                    if ([int]::TryParse($val,[ref]$parsed)) {
                        $encPct = $parsed
                        if ($parsed -ge 100) { $encLabel = 'Encrypted' } else { $encLabel = 'Encrypting' }
                    }
                } else {
                    # try to detect 'Protection On'
                    $prot = $mb | Select-String -Pattern 'Protection Status' -SimpleMatch
                    if ($prot -and $prot -match 'On') { $encLabel = 'Encrypted' }
                }
            }
        }
    } catch {
        # ignore errors and leave defaults
    }

    [PSCustomObject]@{
        Drive      = $_.DeviceID
        Volume     = $_.VolumeName
        SizeBytes  = $size
        UsedBytes  = $used
        FreeBytes  = $free
        Size       = Convert-Bytes $size
        Used       = Convert-Bytes $used
        Free       = Convert-Bytes $free
        PctUsed    = $pctUsed
        PctFree    = $pctFree
        Encrypted  = $encLabel
        PctEncrypted = $encPct
    }
} | Sort-Object Drive

# --- GUI Mode: show drives and allow expandable folder details
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# Capture the drive list into a script-scoped variable so handlers can update it
$script:driveList = @()
function Refresh-DriveList {
    # reset the drive list safely (use a fresh array)
    $script:driveList = @()
    Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=2 OR DriveType=3" |
        Select-Object DeviceID, VolumeName, @{Name='SizeBytes';Expression={[int64]$_.Size}}, @{Name='FreeBytes';Expression={[int64]$_.FreeSpace}} |
        ForEach-Object {
            $size = $_.SizeBytes
            $free = $_.FreeBytes
            $used = $size - $free
            if ($size -gt 0) {
                $pctUsed = ($used / $size) * 100
                $pctFree = ($free / $size) * 100
            } else {
                $pctUsed = 0
                $pctFree = 0
            }
            $script:driveList += [PSCustomObject]@{
                Drive = $_.DeviceID
                Volume = $_.VolumeName
                Size = Convert-Bytes $size
                Used = Convert-Bytes $used
                Free = Convert-Bytes $free
                PctUsed = ('{0:N2}%' -f $pctUsed)
                PctFree = ('{0:N2}%' -f $pctFree)
            }
        }
}

# initial drive enumeration
Refresh-DriveList

function Show-Gui {
    # shared cache used during scanning
    $script:folderSizes = @{}

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Dsize - Disk sizes at a glance'
    $form.Size = New-Object System.Drawing.Size(900,600)
    $form.StartPosition = 'CenterScreen'
    # Apply blue theme to the form
    $form.BackColor = [System.Drawing.Color]::LightSteelBlue
    $form.ForeColor = [System.Drawing.Color]::DarkBlue

    $lv = New-Object System.Windows.Forms.ListView
    $lv.View = 'Details'
    $lv.FullRowSelect = $true
    $lv.Width = 360
    $lv.Height = 480
    $lv.Location = New-Object System.Drawing.Point(10,10)
    $lv.Columns.Add('Drive',80) | Out-Null
    $lv.Columns.Add('Volume',140) | Out-Null
    $lv.Columns.Add('Size',80) | Out-Null
    $lv.Columns.Add('Used',80) | Out-Null
    $lv.Columns.Add('Free',80) | Out-Null
    $lv.Columns.Add('Used %',70) | Out-Null
    $lv.Columns.Add('Free %',70) | Out-Null

    function Populate-Drives {
        $lv.Items.Clear()
        foreach ($d in $script:driveList) {
            $item = New-Object System.Windows.Forms.ListViewItem($d.Drive)
            $item.SubItems.Add($d.Volume) | Out-Null
            $item.SubItems.Add($d.Size) | Out-Null
            $item.SubItems.Add($d.Used) | Out-Null
            $item.SubItems.Add($d.Free) | Out-Null
            $item.SubItems.Add($d.PctUsed) | Out-Null
            $item.SubItems.Add($d.PctFree) | Out-Null
            $lv.Items.Add($item) | Out-Null
        }
    }

    Populate-Drives

    $tree = New-Object System.Windows.Forms.TreeView
    $tree.Width = 480
    $tree.Height = 480
    $tree.Location = New-Object System.Drawing.Point(380,10)
    $tree.HideSelection = $false

    $btnScan = New-Object System.Windows.Forms.Button
    $btnScan.Text = 'Scan Selected Drive'
    $btnScan.Width = 160
    $btnScan.Location = New-Object System.Drawing.Point(10,500)

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = 'Export CSV'
    $btnExport.Width = 120
    $btnExport.Location = New-Object System.Drawing.Point(180,500)

    $btnRescan = New-Object System.Windows.Forms.Button
    $btnRescan.Text = 'Rescan Drives'
    $btnRescan.Width = 120
    $btnRescan.Location = New-Object System.Drawing.Point(310,500)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'
    $btnClose.Width = 80
    $btnClose.Location = New-Object System.Drawing.Point(760,500)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.AutoSize = $true
    $lblStatus.Location = New-Object System.Drawing.Point(320,505)
    $lblStatus.Text = ''

    # Style controls with blue theme accents
    $lv.BackColor = [System.Drawing.Color]::AliceBlue
    $lv.ForeColor = [System.Drawing.Color]::Black

    $tree.BackColor = [System.Drawing.Color]::AliceBlue
    $tree.ForeColor = [System.Drawing.Color]::Black

    $btnScan.BackColor = [System.Drawing.Color]::SteelBlue
    $btnScan.ForeColor = [System.Drawing.Color]::White
    $btnExport.BackColor = [System.Drawing.Color]::SteelBlue
    $btnExport.ForeColor = [System.Drawing.Color]::White
    $btnClose.BackColor = [System.Drawing.Color]::SteelBlue
    $btnClose.ForeColor = [System.Drawing.Color]::White

    $lblStatus.ForeColor = [System.Drawing.Color]::DarkBlue

    $form.Controls.AddRange(@($lv, $tree, $btnScan, $btnExport, $btnRescan, $btnClose, $lblStatus))

    # Helper: build folder-size cache for a root path
    function Build-FolderCache([string]$rootPath) {
        $script:folderSizes.Clear()
        $lblStatus.Text = "Scanning files on $rootPath ..."
        $form.Refresh()
        try {
            $allFiles = Get-ChildItem -Path $rootPath -File -Recurse -Force -ErrorAction SilentlyContinue
        } catch {
            $allFiles = @()
        }
        $total = 0
        if ($allFiles) { $total = $allFiles.Count }
        $count = 0
        if ($total -gt 0) {
            foreach ($file in $allFiles) {
                $count++
                if ($count % 200 -eq 0 -or $count -eq $total) { $lblStatus.Text = "Scanning files: $count / $total"; $form.Refresh() }
                $current = $file.DirectoryName
                while ($current -and $current.StartsWith($rootPath, [System.StringComparison]::OrdinalIgnoreCase)) {
                    if (-not $script:folderSizes.ContainsKey($current)) { $script:folderSizes[$current] = 0 }
                    $script:folderSizes[$current] += $file.Length
                    $parent = Split-Path $current -Parent
                    if ([string]::IsNullOrEmpty($parent) -or $parent -eq $current) { break }
                    $current = $parent
                }
            }
        }
        $lblStatus.Text = "Scan complete"
    }

    # Helper: get immediate child folders from cache
    function Get-ChildFolderSizesCached([string]$Path) {
        try {
            Get-ChildItem -Path $Path -Directory -Force -ErrorAction SilentlyContinue |
                ForEach-Object {
                    $size = 0
                    if ($script:folderSizes.ContainsKey($_.FullName)) { $size = $script:folderSizes[$_.FullName] }
                    [PSCustomObject]@{ Name = $_.Name; FullName = $_.FullName; Size = $size }
                }
        } catch {
            @()
        }
    }

    # When a node expands, lazily add its children from the cache
    $tree.Add_BeforeExpand({
        param($s,$e)
        $node = $e.Node
        if ($node.Nodes.Count -eq 1 -and $node.Nodes[0].Text -eq 'Loading...') {
            $node.Nodes.Clear()
            $children = Get-ChildFolderSizesCached $node.Tag | Where-Object { $_.Size -gt 0 } | Sort-Object Size -Descending
            foreach ($c in $children) {
                $n = New-Object System.Windows.Forms.TreeNode(("{0} -> {1}" -f $c.Name, (Convert-Bytes $c.Size)))
                $n.Tag = $c.FullName
                # add a placeholder child so the node is expandable
                $n.Nodes.Add((New-Object System.Windows.Forms.TreeNode('Loading...'))) | Out-Null
                $node.Nodes.Add($n) | Out-Null
            }
        }
    })

    # Populate root nodes (empty until scanned)
    function Populate-TreeRoot([string]$rootPath, [int]$topN=5) {
        $tree.Nodes.Clear()
        $rootNode = New-Object System.Windows.Forms.TreeNode($rootPath)
        $rootNode.Tag = $rootPath
        # top children
        $children = Get-ChildFolderSizesCached $rootPath | Where-Object { $_.Size -gt 0 } | Sort-Object Size -Descending | Select-Object -First $topN
        foreach ($c in $children) {
            $n = New-Object System.Windows.Forms.TreeNode(("{0} -> {1}" -f $c.Name, (Convert-Bytes $c.Size)))
            $n.Tag = $c.FullName
            $n.Nodes.Add((New-Object System.Windows.Forms.TreeNode('Loading...'))) | Out-Null
            $rootNode.Nodes.Add($n) | Out-Null
        }
        $tree.Nodes.Add($rootNode) | Out-Null
        $rootNode.Expand()
    }

    # Rescan button click
    $btnRescan.Add_Click({
        $btnRescan.Enabled = $false
        $lblStatus.Text = 'Rescanning drives...'
        $form.Refresh()
        try {
            Refresh-DriveList
            Populate-Drives
            $tree.Nodes.Clear()
            if ($script:folderSizes) { $script:folderSizes.Clear() }
            $lblStatus.Text = 'Rescan complete'
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Rescan failed: $_","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        } finally {
            $btnRescan.Enabled = $true
        }
    })

    # Scan button click
    $btnScan.Add_Click({
        if ($lv.SelectedItems.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('Select a drive first','No drive selected',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information); return }
        $drive = $lv.SelectedItems[0].Text
        $rootPath = "$drive\"
        $topN = 5
        # ask user for topN using InputBox-style prompt
        $input = [Microsoft.VisualBasic.Interaction]::InputBox('How many top folders to show? (1-20)','Top folders','5')
        if ($input -and [int]::TryParse($input,[ref]$null)) { $topN = [int]$input }
        $btnScan.Enabled = $false
        $btnExport.Enabled = $false
        try {
            Build-FolderCache $rootPath
            Populate-TreeRoot $rootPath $topN
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to scan: $_","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        } finally {
            $btnScan.Enabled = $true
            $btnExport.Enabled = $true
        }
    })

    # Export button click
    $btnExport.Add_Click({
        if (-not $script:folderSizes -or $script:folderSizes.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('No scan data to export. Please scan a drive first.','No data',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information); return }
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
        $sfd.FileName = ("Dsize_export_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmm'))
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $exportList = $script:folderSizes.GetEnumerator() | ForEach-Object {[PSCustomObject]@{ FullName = $_.Key; SizeBytes = $_.Value; Size = Convert-Bytes $_.Value }}
                $exportList | Sort-Object SizeBytes -Descending | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8
                [System.Windows.Forms.MessageBox]::Show("Exported to $($sfd.FileName)", 'Export', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to export: $_", 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })

    $btnClose.Add_Click({ $form.Close() })

    # Show form
    [void] [System.Windows.Forms.Application]::Run($form)
}

# If running in a non-interactive host (no GUI), fall back to console table
if ($Host.Name -match 'ConsoleHost' -and [System.Environment]::UserInteractive) {
    Show-Gui
} else {
    # Some hosts (ISE) support GUI; try to show it, otherwise keep console output
    try {
        Show-Gui
    } catch {
        Write-Host "Unable to show GUI; running console output instead." -ForegroundColor Yellow
        # Re-run the table output that was previously generated
        $driveList | Format-Table -AutoSize
    }
}
