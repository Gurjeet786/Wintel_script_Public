
<# 
Complete folder inventory (long-path safe) that captures:
- Folder Owner, full NTFS ACLs, inheritance details
- Share metadata and Share permissions for any folder under a share
- Full normal UNC and extended UNC path
- Adds a row for inaccessible folders with AccessStatus = "Unaccessible"
- Forces traversal into subfolders using BFS and a visited set
#>

# ======================== Config ========================
$serverName   = $env:COMPUTERNAME
$dateStamp    = (Get-Date -Format "yyyyMMdd_HHmmss")
$reportFolder = "C:\FSS_report"

# Create report folder if not exists
if (-not (Test-Path -LiteralPath $reportFolder)) {
    New-Item -Path $reportFolder -ItemType Directory | Out-Null
}

# Output files
$csvPath  = "$reportFolder\$serverName`_$dateStamp.csv"
$xlsxPath = "$reportFolder\$serverName`_$dateStamp.xlsx"

# Drives to scan (exclude C: per earlier requirement; change if needed)
$drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Name -ne 'C' }

# ======================== Helpers ========================

function ConvertTo-LongPath {
    param([Parameter(Mandatory)][string]$Path)
    if ($Path -like '\\*') { return ('\\?\UNC' + $Path.Substring(1)) }
    else { return ('\\?\' + $Path) }
}

function ConvertFrom-LongPath {
    param([Parameter(Mandatory)][string]$Path)
    if ($Path -like '\\?\UNC\*') { return ('\' + $Path.Substring(7)) }
    elseif ($Path -like '\\?\*') { return ($Path.Substring(4)) }
    else { return $Path }
}

function Normalize-TrailingSlash {
    param([Parameter(Mandatory)][string]$Path)
    if ($Path.EndsWith('\')) { $Path } else { $Path + '\' }
}

# Preload SMB shares and their permissions once
$allShares = @()
$shareAccessMap = @{} # Key: ShareName -> Value: list of access rows
try {
    $allShares = Get-SmbShare -ErrorAction SilentlyContinue | Where-Object { $_.Path -and $_.Name }
    foreach ($s in $allShares) {
        try {
            $acc = Get-SmbShareAccess -Name $s.Name -ErrorAction SilentlyContinue
            $shareAccessMap[$s.Name] = $acc
        } catch {
            $shareAccessMap[$s.Name] = @()
        }
    }
} catch {
    Write-Host "Warning: Unable to enumerate SMB shares on $serverName."
}

function Resolve-NetworkPaths {
    param(
        [Parameter(Mandatory)][string]$FullPath,
        [Parameter(Mandatory)][string]$ServerName,
        [Parameter(Mandatory)][object[]]$Shares
    )

    if ($FullPath -like '\\*') {
        $normalUNC   = $FullPath
        $extendedUNC = ConvertTo-LongPath $FullPath
        return [PSCustomObject]@{
            NormalUNC        = $normalUNC
            ExtendedUNC      = $extendedUNC
            Share            = $null
            SharePath        = ''
            ShareName        = ''
            ShareDescription = ''
        }
    }

    $normalizedFull = Normalize-TrailingSlash $FullPath
    $containingShare = $Shares |
        Where-Object {
            $sp = Normalize-TrailingSlash $_.Path
            $normalizedFull.StartsWith($sp, [System.StringComparison]::OrdinalIgnoreCase)
        } |
        Sort-Object { $_.Path.Length } -Descending |
        Select-Object -First 1

    if ($containingShare) {
        $basePath  = Normalize-TrailingSlash $containingShare.Path
        $relative  = $FullPath.Substring($basePath.Length).TrimStart('\')
        $uncRoot   = "\\$ServerName\$($containingShare.Name)"
        $normalUNC = if ($relative) { Join-Path $uncRoot $relative } else { $uncRoot }
        return [PSCustomObject]@{
            NormalUNC        = $normalUNC
            ExtendedUNC      = ConvertTo-LongPath $normalUNC
            Share            = $containingShare
            SharePath        = $containingShare.Path
            ShareName        = $containingShare.Name
            ShareDescription = $containingShare.Description
        }
    }

    # Fallback → administrative share
    $driveLetter = $FullPath.Substring(0,1)
    $relative    = ($FullPath -replace '^[A-Za-z]:\\','')
    $normalUNC   = "\\$ServerName\$driveLetter`$\$relative"
    return [PSCustomObject]@{
        NormalUNC        = $normalUNC
        ExtendedUNC      = ConvertTo-LongPath $normalUNC
        Share            = $null
        SharePath        = ''
        ShareName        = ''
        ShareDescription = ''
    }
}

# Robust child directory enumeration (long-path safe)
function Get-ChildDirectoriesSafe {
    param([Parameter(Mandatory)][string]$LongPath)

    # Prefer .NET enumeration on extended path
    try {
        $dirs = [System.IO.Directory]::EnumerateDirectories($LongPath)
        if ($dirs) { return $dirs }
    } catch {
        # Fallback to PowerShell provider on normal path
        try {
            $normal = ConvertFrom-LongPath $LongPath
            $items = Get-ChildItem -LiteralPath $normal -Directory -Force -ErrorAction SilentlyContinue
            if ($items) { return ($items | ForEach-Object { $_.FullName }) }
        } catch { }
    }
    return @()
}

# Common row builder
function Add-ResultRow {
    param(
        [string]$RootFolderPath,
        [string]$SubFolderPath,
        [string]$FullNetworkPath,
        [string]$ExtendedNetworkPath,
        [string]$Owner,
        [string]$ACL_UserGroup,
        [string]$ACL_PermissionLevel,
        [string]$ACL_AccessType,
        [string]$InheritFrom,
        [string]$InheritanceStatus,
        [string]$SharePath,
        [string]$ShareName,
        [string]$ShareDescription,
        [string]$Share_UserGroup = "",
        [string]$Share_PermissionLevel = "",
        [bool]  $IsReparsePoint = $false,
        [string]$LinkTarget = "",
        [string]$AccessStatus = "Accessible"
    )
    $folderName = Split-Path -Path $SubFolderPath -Leaf
    $script:results.Add([PSCustomObject]@{
        HostName              = $serverName
        FolderName            = $folderName
        RootFolderPath        = $RootFolderPath
        SubFolderPath         = $SubFolderPath
        FullNetworkPath       = $FullNetworkPath
        ExtendedNetworkPath   = $ExtendedNetworkPath
        Owner                 = $Owner
        ACL_UserGroup         = $ACL_UserGroup
        ACL_PermissionLevel   = $ACL_PermissionLevel
        ACL_AccessType        = $ACL_AccessType
        InheritFrom           = $InheritFrom
        InheritanceStatus     = $InheritanceStatus
        SharePath             = $SharePath
        ShareName             = $ShareName
        ShareDescription      = $ShareDescription
        Share_UserGroup       = $Share_UserGroup
        Share_PermissionLevel = $Share_PermissionLevel
        IsReparsePoint        = $IsReparsePoint
        LinkTarget            = $LinkTarget
        AccessStatus          = $AccessStatus
    })
}

# ======================== Main ========================

$results = New-Object System.Collections.Generic.List[object]

foreach ($drive in $drives) {
    Write-Host "Scanning drive: $($drive.Name):\"

    # Drive readiness check
    $driveInfo = New-Object System.IO.DriveInfo($drive.Name)
    if (-not $driveInfo.IsReady) {
        $resolvedRoot = Resolve-NetworkPaths -FullPath $drive.Root -ServerName $serverName -Shares $allShares
        Add-ResultRow -RootFolderPath $drive.Root -SubFolderPath $drive.Root `
            -FullNetworkPath $resolvedRoot.NormalUNC -ExtendedNetworkPath $resolvedRoot.ExtendedUNC `
            -Owner "Drive not ready" -AccessStatus "Unaccessible" -SharePath $resolvedRoot.SharePath `
            -ShareName $resolvedRoot.ShareName -ShareDescription $resolvedRoot.ShareDescription
        Write-Host "Drive $($drive.Name): not ready. Recorded and continuing."
        continue
    }

    # BFS queue and visited set to guarantee traversal
    $queue   = New-Object System.Collections.Generic.Queue[string]
    $visited = New-Object System.Collections.Generic.HashSet[string]

    $queue.Enqueue($drive.Root)
    $visited.Add($drive.Root.ToLowerInvariant()) | Out-Null

    $processed = 0
    while ($queue.Count -gt 0) {
        $currentNormal = $queue.Dequeue()
        $currentLong   = ConvertTo-LongPath $currentNormal
        $resolved      = Resolve-NetworkPaths -FullPath $currentNormal -ServerName $serverName -Shares $allShares

        # Try ACL for the current folder
        $owner = ""
        $acl   = $null
        $aclSucceeded = $false
        try {
            $acl   = Get-Acl -LiteralPath $currentLong -ErrorAction Stop
            $owner = $acl.Owner
            $aclSucceeded = $true

            # Record ALL ACEs
            foreach ($entry in $acl.Access) {
                Add-ResultRow -RootFolderPath $drive.Root -SubFolderPath $currentNormal `
                    -FullNetworkPath $resolved.NormalUNC -ExtendedNetworkPath $resolved.ExtendedUNC `
                    -Owner $owner `
                    -ACL_UserGroup $entry.IdentityReference `
                    -ACL_PermissionLevel $entry.FileSystemRights `
                    -ACL_AccessType $entry.AccessControlType `
                    -InheritFrom (if ($entry.IsInherited) { "Inherited" } else { "Direct" }) `
                    -InheritanceStatus (if ($acl.AreAccessRulesProtected) { "Disabled" } else { "Enabled" }) `
                    -SharePath $resolved.SharePath -ShareName $resolved.ShareName -ShareDescription $resolved.ShareDescription `
                    -AccessStatus "Accessible"
            }

            # If a folder has no ACEs (possible), still write one row so owner/inheritance are captured
            if (-not $acl.Access -or $acl.Access.Count -eq 0) {
                Add-ResultRow -RootFolderPath $drive.Root -SubFolderPath $currentNormal `
                    -FullNetworkPath $resolved.NormalUNC -ExtendedNetworkPath $resolved.ExtendedUNC `
                    -Owner $owner `
                    -ACL_UserGroup "" -ACL_PermissionLevel "" -ACL_AccessType "" `
                    -InheritFrom "" -InheritanceStatus (if ($acl.AreAccessRulesProtected) { "Disabled" } else { "Enabled" }) `
                    -SharePath $resolved.SharePath -ShareName $resolved.ShareName -ShareDescription $resolved.ShareDescription `
                    -AccessStatus "Accessible"
            }
        } catch {
            # Record the folder as Unaccessible
            Add-ResultRow -RootFolderPath $drive.Root -SubFolderPath $currentNormal `
                -FullNetworkPath $resolved.NormalUNC -ExtendedNetworkPath $resolved.ExtendedUNC `
                -Owner "Access Denied / Not Found" `
                -ACL_UserGroup "" -ACL_PermissionLevel "" -ACL_AccessType "" `
                -InheritFrom "" -InheritanceStatus "" `
                -SharePath $resolved.SharePath -ShareName $resolved.ShareName -ShareDescription $resolved.ShareDescription `
                -AccessStatus "Unaccessible"
            Write-Host "ACL read failed: $currentNormal"
        }

        # Add Share permission rows for EVERY folder within a share
        if ($resolved.Share) {
            $sharePerms = @()
            if ($shareAccessMap.ContainsKey($resolved.Share.Name)) {
                $sharePerms = $shareAccessMap[$resolved.Share.Name]
            }
            foreach ($perm in $sharePerms) {
                Add-ResultRow -RootFolderPath $drive.Root -SubFolderPath $currentNormal `
                    -FullNetworkPath $resolved.NormalUNC -ExtendedNetworkPath $resolved.ExtendedUNC `
                    -Owner $owner `
                    -ACL_UserGroup "" -ACL_PermissionLevel "" -ACL_AccessType "" `
                    -InheritFrom "" -InheritanceStatus "N/A" `
                    -SharePath $resolved.SharePath -ShareName $resolved.ShareName -ShareDescription $resolved.ShareDescription `
                    -Share_UserGroup $perm.AccountName -Share_PermissionLevel $perm.AccessRight `
                    -AccessStatus (if ($aclSucceeded) { "Accessible" } else { "Unaccessible" })
            }
            # If no share permissions returned, still include one row to capture share metadata
            if (-not $sharePerms -or $sharePerms.Count -eq 0) {
                Add-ResultRow -RootFolderPath $drive.Root -SubFolderPath $currentNormal `
                    -FullNetworkPath $resolved.NormalUNC -ExtendedNetworkPath $resolved.ExtendedUNC `
                    -Owner $owner `
                    -SharePath $resolved.SharePath -ShareName $resolved.ShareName -ShareDescription $resolved.ShareDescription `
                    -InheritanceStatus "N/A" -AccessStatus (if ($aclSucceeded) { "Accessible" } else { "Unaccessible" })
            }
        }

        # Enumerate children (long-path safe)
        $childDirs = Get-ChildDirectoriesSafe -LongPath $currentLong

        if (-not $childDirs -or $childDirs.Count -eq 0) {
            # Still ensure we mark any enumeration issue as Unaccessible for this folder
            # (but we already added current folder rows above)
            # No enqueue needed here if there are none
        } else {
            foreach ($childLong in $childDirs) {
                $childNormal = ConvertFrom-LongPath $childLong

                # Prevent infinite loops / duplicates
                $key = $childNormal.ToLowerInvariant()
                if (-not $visited.Contains($key)) {
                    $visited.Add($key) | Out-Null
                    # Enqueue to guarantee ACL capture for the child itself later
                    $queue.Enqueue($childNormal)
                }

                # Record minimal child presence row immediately (so it's visible even if later ACL fails)
                $childResolved = Resolve-NetworkPaths -FullPath $childNormal -ServerName $serverName -Shares $allShares
                Add-ResultRow -RootFolderPath $drive.Root -SubFolderPath $childNormal `
                    -FullNetworkPath $childResolved.NormalUNC -ExtendedNetworkPath $childResolved.ExtendedUNC `
                    -Owner "" -ACL_UserGroup "" -ACL_PermissionLevel "" -ACL_AccessType "" `
                    -InheritFrom "" -InheritanceStatus "" `
                    -SharePath $childResolved.SharePath -ShareName $childResolved.ShareName -ShareDescription $childResolved.ShareDescription `
                    -AccessStatus "Accessible"
            }
        }

        $processed++
        if ($processed % 1000 -eq 0) {
            Write-Progress -Activity "Scanning folders" -Status "$processed processed on $($drive.Name):\" -PercentComplete 0
        }
    }
}

# Export to CSV (UTF-8)
$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "CSV Report exported to $csvPath"

# Optional: Export to Excel (requires ImportExcel)
# Install-Module -Name ImportExcel -Force
# $results | Export-Excel -Path $xlsxPath -AutoSize
# Write-Host "Excel Report exported to $xlsxPath"
