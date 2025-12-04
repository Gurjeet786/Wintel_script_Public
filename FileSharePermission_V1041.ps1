
<# 
Long-path safe folder & ACL inventory with complete UNC capture.
- Logs every folder and subfolder
- Records a row for inaccessible folders with AccessStatus = "Unaccessible"
- Stores FolderName, SubFolderPath, FullNetworkPath, ExtendedNetworkPath
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

# Drives to scan (exclude C: per original; adjust if needed)
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

# Preload SMB shares once
$allShares = @()
try {
    $allShares = Get-SmbShare -ErrorAction SilentlyContinue | Where-Object { $_.Path -and $_.Name }
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

# Child directory enumeration that survives root issues and long paths
function Get-ChildDirectoriesSafe {
    param([Parameter(Mandatory)][string]$LongPath)

    # First try .NET extended-length enumeration
    try {
        $dirs = [System.IO.Directory]::EnumerateDirectories($LongPath)
        if ($dirs) { return $dirs }
    } catch {
        # Fall back to PowerShell provider with normal path
        try {
            $normal = ConvertFrom-LongPath $LongPath
            $items = Get-ChildItem -LiteralPath $normal -Directory -Force -ErrorAction SilentlyContinue
            if ($items) { return ($items | ForEach-Object { $_.FullName }) }
        } catch { }
    }
    return @()
}

# Common row builder to reduce duplication
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
        AccessStatus          = $AccessStatus   # "Accessible" or "Unaccessible"
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
            -Owner "Drive not ready" -ACL_UserGroup "" -ACL_PermissionLevel "" -ACL_AccessType "" `
            -InheritFrom "" -InheritanceStatus "" -SharePath $resolvedRoot.SharePath `
            -ShareName $resolvedRoot.ShareName -ShareDescription $resolvedRoot.ShareDescription `
            -AccessStatus "Unaccessible"
        Write-Host "Drive $($drive.Name): is not ready. Recorded and continuing."
        continue
    }

    # BFS queue
    $queue = New-Object System.Collections.Generic.Queue[string]
    $queue.Enqueue($drive.Root)

    $processed = 0
    while ($queue.Count -gt 0) {
        $currentNormal = $queue.Dequeue()
        $currentLong   = ConvertTo-LongPath $currentNormal

        # Resolve network paths
        $resolved = Resolve-NetworkPaths -FullPath $currentNormal -ServerName $serverName -Shares $allShares

        # Try ACL
        $owner = ""
        $aclSucceeded = $false
        try {
            $acl   = Get-Acl -LiteralPath $currentLong -ErrorAction Stop
            $owner = $acl.Owner
            $aclSucceeded = $true

            foreach ($entry in $acl.Access) {
                Add-ResultRow -RootFolderPath $drive.Root -SubFolderPath $currentNormal `
                    -FullNetworkPath $resolved.NormalUNC -ExtendedNetworkPath $resolved.ExtendedUNC `
                    -Owner $owner -ACL_UserGroup $entry.IdentityReference -ACL_PermissionLevel $entry.FileSystemRights `
                    -ACL_AccessType $entry.AccessControlType -InheritFrom (if ($entry.IsInherited) { "Inherited" } else { "Direct" }) `
                    -InheritanceStatus (if ($acl.AreAccessRulesProtected) { "Disabled" } else { "Enabled" }) `
                    -SharePath $resolved.SharePath -ShareName $resolved.ShareName -ShareDescription $resolved.ShareDescription `
                    -AccessStatus "Accessible"
            }
        } catch {
            # Log inaccessible ACL
            Add-ResultRow -RootFolderPath $drive.Root -SubFolderPath $currentNormal `
                -FullNetworkPath $resolved.NormalUNC -ExtendedNetworkPath $resolved.ExtendedUNC `
                -Owner "Access Denied / Not Found" -ACL_UserGroup "" -ACL_PermissionLevel "" -ACL_AccessType "" `
                -InheritFrom "" -InheritanceStatus "" -SharePath $resolved.SharePath -ShareName $resolved.ShareName `
                -ShareDescription $resolved.ShareDescription -AccessStatus "Unaccessible"
            Write-Host "Skipped ACL (Access/Existence): $currentNormal"
        }

        # Share permissions (if under a share)
        if ($resolved.Share -and $aclSucceeded) {
            try {
                $sharePerms = Get-SmbShareAccess -Name $resolved.Share.Name -ErrorAction Stop
                foreach ($perm in $sharePerms) {
                    Add-ResultRow -RootFolderPath $drive.Root -SubFolderPath $currentNormal `
                        -FullNetworkPath $resolved.NormalUNC -ExtendedNetworkPath $resolved.ExtendedUNC `
                        -Owner $owner -ACL_UserGroup "" -ACL_PermissionLevel "" -ACL_AccessType "" `
                        -InheritFrom "" -InheritanceStatus "N/A" -SharePath $resolved.SharePath `
                        -ShareName $resolved.ShareName -ShareDescription $resolved.ShareDescription `
                        -Share_UserGroup $perm.AccountName -Share_PermissionLevel $perm.AccessRight `
                        -AccessStatus "Accessible"
                }
            } catch {
                # If share permission retrieval itself fails, log as Unaccessible
                Add-ResultRow -RootFolderPath $drive.Root -SubFolderPath $currentNormal `
                    -FullNetworkPath $resolved.NormalUNC -ExtendedNetworkPath $resolved.ExtendedUNC `
                    -Owner $owner -SharePath $resolved.SharePath -ShareName $resolved.ShareName `
                    -ShareDescription $resolved.ShareDescription -AccessStatus "Unaccessible"
                Write-Host "Share permission retrieval error: $($resolved.Share.Name)"
            }
        }

        # Enumerate child directories with robust, long-path safe method
        $childDirs = @()
        $childEnumSucceeded = $true
        try {
            $childDirs = Get-ChildDirectoriesSafe -LongPath $currentLong
        } catch {
            $childEnumSucceeded = $false
            $childDirs = @()
        }

        if (-not $childDirs) {
            # Explicitly log child enumeration failure for this folder (root or any)
            Add-ResultRow -RootFolderPath $drive.Root -SubFolderPath $currentNormal `
                -FullNetworkPath $resolved.NormalUNC -ExtendedNetworkPath $resolved.ExtendedUNC `
                -Owner "" -AccessStatus "Unaccessible"
            Write-Host "Child enumeration error at: $currentNormal"
        }

        foreach ($childLongPath in $childDirs) {
            $childNormal = ConvertFrom-LongPath $childLongPath

            # Get attributes safely (to flag reparse points)
            $isReparse = $false
            $linkTarget = ""
            try {
                $childItem = Get-Item -LiteralPath $childLongPath -Force -ErrorAction Stop
                $isReparse = ($childItem.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0
                if ($isReparse -and $childItem.PSObject.Properties['LinkTarget']) {
                    $linkTarget = ($childItem.LinkTarget -join ';')
                }
            } catch {
                # Even if we can't read attributes, still record child as Unaccessible
                $childResolved = Resolve-NetworkPaths -FullPath $childNormal -ServerName $serverName -Shares $allShares
                Add-ResultRow -RootFolderPath $drive.Root -SubFolderPath $childNormal `
                    -FullNetworkPath $childResolved.NormalUNC -ExtendedNetworkPath $childResolved.ExtendedUNC `
                    -Owner "" -AccessStatus "Unaccessible"
                continue
            }

            # Resolve network paths and record the child (so nothing is skipped)
            $childResolved = Resolve-NetworkPaths -FullPath $childNormal -ServerName $serverName -Shares $allShares
            Add-ResultRow -RootFolderPath $drive.Root -SubFolderPath $childNormal `
                -FullNetworkPath $childResolved.NormalUNC -ExtendedNetworkPath $childResolved.ExtendedUNC `
                -Owner "" -IsReparsePoint $isReparse -LinkTarget $linkTarget `
                -AccessStatus "Accessible"  # child directory exists; ACL may be processed later when dequeued

            # Traverse only non-reparse children to avoid cycles; still recorded above
            if (-not $isReparse) {
                $queue.Enqueue($childNormal)
            }
        }

        $processed++
        if ($processed % 500 -eq 0) {
            Write-Progress -Activity "Scanning folders" -Status "$processed processed on $($drive.Name):\" -PercentComplete 0
        }
    }
}

# Export to CSV
$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "CSV Report exported to $csvPath"

# Optional: Export to Excel (requires ImportExcel)
# Install-Module -Name ImportExcel -Force
# $results | Export-Excel -Path $xlsxPath -AutoSize
# Write-Host "Excel Report exported to $xlsxPath"
