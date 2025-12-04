# PowerShell script: list drives with size, used, free and percentages
# This script has been created in partnership with my AI Buddy
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
} | Sort-Object Drive |
Format-Table -Property `
    @{Name='Drive';Expression={$_.Drive}}, `
    @{Name='Volume';Expression={$_.Volume}}, `
    @{Name='Size';Expression={$_.Size};Align='Right'}, `
    @{Name='Used';Expression={$_.Used};Align='Right'}, `
    @{Name='Free';Expression={$_.Free};Align='Right'}, `
    @{Name='Encrypted';Expression={$_.Encrypted}}, `
    @{Name='%Enc';Expression={ if ($_.PctEncrypted -ne $null) { ('{0:N0}%' -f $_.PctEncrypted) } else { '' } };Align='Right'}, `
    @{Name='%Used';Expression={('{0:N2}%' -f $_.PctUsed)};Align='Right'}, `
    @{Name='%Free';Expression={('{0:N2}%' -f $_.PctFree)};Align='Right'} `
    -AutoSize

    while ($true) {
        # Prompt user
        $see = Read-Host "Would you like to see what's taking the most space on a drive? (Y/N)"
        if ($see -notmatch '^(y|Y)') { break }

        # Get available drives
        $drives = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=2 OR DriveType=3" |
            Select-Object -ExpandProperty DeviceID
        $drivePrompt = Read-Host "Enter drive (e.g. C:) or choose from: $($drives -join ', ')"
        if (-not $drivePrompt) { Write-Host "No drive entered. Exiting."; break }
        $driveLetter = ($drivePrompt.Trim()).TrimEnd('\')
        if ($driveLetter -notmatch '^[A-Za-z]:$') { Write-Host "Invalid drive format. Use like C:. Exiting."; break }
        $rootPath = "$driveLetter\"

        # Ask how many top folders
        do {
            $n = Read-Host "How many of the top largest folders would you like to see? (1-10)"
        } until ([int]::TryParse($n,[ref]$null) -and [int]$n -ge 1 -and [int]$n -le 10)
        $topN = [int]$n

        # Colors to cycle by depth
        $colors = 'Cyan','Yellow','Green','Magenta','White','DarkCyan','DarkYellow','DarkGreen','Red','Blue'

        # Build a folder-size cache with a single recursive file scan (much faster than per-folder recursion)
        Write-Host "Scanning files on ${rootPath} ... (building folder-size cache)" -ForegroundColor DarkGray

        # Collect files into an array to compute total for progress reporting
        $allFiles = Get-ChildItem -Path $rootPath -File -Recurse -Force -ErrorAction SilentlyContinue
        $total = 0
        if ($allFiles) { $total = $allFiles.Count }
        $count = 0

        # Hashtable mapping folder full path -> cumulative size (bytes)
        $folderSizes = @{}

        if ($total -gt 0) {
            foreach ($file in $allFiles) {
                $count++
                if ($count % 200 -eq 0 -or $count -eq $total) {
                    $percent = [int](($count / $total) * 100)
                    Write-Progress -Activity "Scanning files" -Status "Processing $count of $total files" -PercentComplete $percent
                }

                $current = $file.DirectoryName
                while ($current -and $current.StartsWith($rootPath, [System.StringComparison]::OrdinalIgnoreCase)) {
                    if (-not $folderSizes.ContainsKey($current)) { $folderSizes[$current] = 0 }
                    $folderSizes[$current] += $file.Length
                    $parent = Split-Path $current -Parent
                    if ([string]::IsNullOrEmpty($parent) -or $parent -eq $current) { break }
                    $current = $parent
                }
            }
            Write-Progress -Activity "Scanning files" -Completed
        } else {
            Write-Progress -Activity "Scanning files" -Completed
        }

        # Helper: get immediate child folders with their sizes from the cache
        function Get-ChildFolderSizesCached([string]$Path) {
            Get-ChildItem -Path $Path -Directory -Force -ErrorAction SilentlyContinue |
                ForEach-Object {
                    $size = 0
                    if ($folderSizes.ContainsKey($_.FullName)) { $size = $folderSizes[$_.FullName] }
                    [PSCustomObject]@{ Name = $_.Name; FullName = $_.FullName; Size = $size }
                }
        }

        # Recursive display that uses the cached sizes
        function Show-FolderTreeCached([string]$Path, [int]$Depth=0) {
            $indent = ' ' * ($Depth * 4)
            $color = $colors[$Depth % $colors.Count]
            $size = 0
            if ($folderSizes.ContainsKey($Path)) { $size = $folderSizes[$Path] }
            Write-Host ("{0}{1} -> {2}" -f $indent, (Split-Path $Path -Leaf), (Convert-Bytes $size)) -ForegroundColor $color

            # Show only the top 3 largest child folders for each node to keep output concise
            $children = Get-ChildFolderSizesCached $Path | Where-Object { $_.Size -gt 0 } | Sort-Object Size -Descending | Select-Object -First 3
            foreach ($c in $children) {
                Show-FolderTreeCached $c.FullName ($Depth + 1)
            }
        }

        # Find top N folders at root of the drive using the cache
        $topFolders = Get-ChildFolderSizesCached $rootPath | Sort-Object Size -Descending | Select-Object -First $topN

        if (-not $topFolders) {
            Write-Host "No folders found or unable to enumerate $rootPath" -ForegroundColor Red
        } else {
            Write-Host "Top $topN folders on ${rootPath}:" -ForegroundColor Cyan
            foreach ($f in $topFolders) {
                # show the folder tree for each top folder (color depth starts at 0 for top folder)
                Show-FolderTreeCached $f.FullName 0
                Write-Host ""  # blank line between top folders
            }

            # Offer to export results to CSV
            $export = Read-Host "Would you like to export folder sizes to CSV? (Y/N)"
            if ($export -match '^(y|Y)') {
                $safeDrive = $driveLetter.TrimEnd(':')
                $defaultPath = [IO.Path]::Combine($env:USERPROFILE, 'Desktop', ("Dsize_{0}_{1}.csv" -f $safeDrive, (Get-Date -Format 'yyyyMMdd_HHmm')))
                $path = Read-Host "Enter path to save CSV (default: $defaultPath)"
                if (-not $path) { $path = $defaultPath }
                try {
                    $exportList = $folderSizes.GetEnumerator() | ForEach-Object {[PSCustomObject]@{ FullName = $_.Key; SizeBytes = $_.Value; Size = Convert-Bytes $_.Value }}
                    $exportList | Sort-Object SizeBytes -Descending | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
                    Write-Host "Exported to $path" -ForegroundColor Green
                } catch {
                    Write-Host "Failed to export: $_" -ForegroundColor Red
                }
            }
        }

        # Ask whether to scan another drive
        $again = Read-Host "Would you like to scan another drive? (Y/N)"
        if ($again -match '^(y|Y)') {
            Clear-Host
            continue
        } else {
            break
        }
    }