
<# 
Long-path safe folder & ACL inventory with complete UNC capture.
- Enumerates every folder and logs even when access errors occur.
- Uses extended-length literal paths (\\?\ and \\?\UNC\...) to avoid 260-char limits.
- Records Normal UNC and Extended UNC for each folder.
- Includes reparse points in the report but does not traverse inside them.
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

# Drives to scan (exclude C: per your original; adjust if needed)
$drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Name -ne 'C' }

# ======================== Helpers ========================

function ConvertTo-LongPath {
    param([Parameter(Mandatory)][string]$Path)
    # Return extended-length literal path for local or UNC paths
    if ($Path -like '\\*') {
        # \\server\share\path -> \\?\UNC\server\share\path
        return ('\\?\UNC' + $Path.Substring(1))
    } else {
        # C:\path -> \\?\C:\path
        return ('\\?\' + $Path)
    }
}

function ConvertFrom-LongPath {
    param([Parameter(Mandatory)][string]$Path)
    # Strip \\?\ or \\?\UNC back to normal paths for display
    if ($Path -like '\\?\UNC\*') {
        return ('\' + $Path.Substring(7))  # 7 = length of "?\UNC"
    } elseif ($Path -like '\\?\*') {
        return ($Path.Substring(4))        # remove "\\?\"
    } else {
        return $Path
    }
}

function Normalize-TrailingSlash {
    param([Parameter(Mandatory)][string]$Path)
    # Ensure a trailing backslash for comparisons
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
        [Parameter(Mandatory)][string]$FullPath,     # normal provider path (not extended)
        [Parameter(Mandatory)][string]$ServerName,
        [Parameter(Mandatory)][object[]]$Shares
    )
    # Already UNC?
    if ($FullPath -like '\\*') {
        $normalUNC  = $FullPath
        $extendedUNC= ConvertTo-LongPath $FullPath
        return [PSCustomObject]@{
            NormalUNC         = $normalUNC
            ExtendedUNC       = $extendedUNC
            Share             = $null
            SharePath         = ''
            ShareName         = ''
            ShareDescription  = ''
        }
    }

    # Find containing share (longest-path prefix match)
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
        $relative  = $FullPath.Substring($basePath.Length)
        $relative  = $relative.TrimStart('\')
        $uncRoot   = "\\$ServerName\$($containingShare.Name)"
        $normalUNC = if ($relative) { Join-Path $uncRoot $relative } else { $uncRoot }
        $extended  = ConvertTo-LongPath $normalUNC
        return [PSCustomObject]@{
            NormalUNC         = $normalUNC
            ExtendedUNC       = $extended
            Share             = $containingShare
            SharePath         = $containingShare.Path
            ShareName         = $containingShare.Name
            ShareDescription  = $containingShare.Description
        }
    }

    # Fallback → Administrative share (e.g., \\Server\D$\...)
    $driveLetter = $FullPath.Substring(0,1)
    $relative    = ($FullPath -replace '^[A-Za-z]:\\','')
    $normalUNC   = "\\$ServerName\$driveLetter`$\$relative"
    $extended    = ConvertTo-LongPath $normalUNC
    return [PSCustomObject]@{
        NormalUNC         = $normalUNC
        ExtendedUNC       = $extended
        Share             = $null
        SharePath         = ''
        ShareName         = ''
        ShareDescription  = ''
    }
}

# ======================== Main ========================

$results = New-Object System.Collections.Generic.List[object]

foreach ($drive in $drives) {
    Write-Host "Scanning drive: $($drive.Name):\"

    # Use BFS queue for reliable traversal (handles long paths & per-folder error handling)
    $queue = New-Object System.Collections.Generic.Queue[string]
    $queue.Enqueue($drive.Root)

    $processed = 0
    while ($queue.Count -gt 0) {
        $currentNormal = $queue.Dequeue()
        $currentLong   = ConvertTo-LongPath $currentNormal

        # Resolve network paths (both normal and extended UNC)
        $resolved = Resolve-NetworkPaths -FullPath $currentNormal -ServerName $serverName -Shares $allShares

        # Try to get ACL information
        $owner = ""
        try {
            $acl = Get-Acl -LiteralPath $currentLong -ErrorAction Stop
            $owner = $acl.Owner

            foreach ($entry in $acl.Access) {
                $results.Add([PSCustomObject]@{
                    HostName              = $serverName
                    RootFolderPath        = $drive.Root
                    SubFolderPath         = $currentNormal
                    FullNetworkPath       = $resolved.NormalUNC
                    ExtendedNetworkPath   = $resolved.ExtendedUNC
                    Owner                 = $owner
                    ACL_UserGroup         = $entry.IdentityReference
                    ACL_PermissionLevel   = $entry.FileSystemRights
                    ACL_AccessType        = $entry.AccessControlType
                    InheritFrom           = if ($entry.IsInherited) { "Inherited" } else { "Direct" }
                    InheritanceStatus     = if ($acl.AreAccessRulesProtected) { "Disabled" } else { "Enabled" }
                    SharePath             = $resolved.SharePath
                    ShareName             = $resolved.ShareName
                    ShareDescription      = $resolved.ShareDescription
                    Share_UserGroup       = ""
                    Share_PermissionLevel = ""
                    IsReparsePoint        = $false
                    LinkTarget            = ""
                })
            }
        } catch {
            # Record folder even if access denied / not found
            $results.Add([PSCustomObject]@{
                HostName              = $serverName
                RootFolderPath        = $drive.Root
                SubFolderPath         = $currentNormal
                FullNetworkPath       = $resolved.NormalUNC
                ExtendedNetworkPath   = $resolved.ExtendedUNC
                Owner                 = "Access Denied / Not Found"
                ACL_UserGroup         = ""
                ACL_PermissionLevel   = ""
                ACL_AccessType        = ""
                InheritFrom           = ""
                InheritanceStatus     = ""
                SharePath             = $resolved.SharePath
                ShareName             = $resolved.ShareName
                ShareDescription      = $resolved.ShareDescription
                Share_UserGroup       = ""
                Share_PermissionLevel = ""
                IsReparsePoint        = $false
                LinkTarget            = ""
            })
            Write-Host "Skipped ACL (Access/Existence): $currentNormal"
        }

        # Share permissions for containers under a share
        if ($resolved.Share) {
            try {
                $sharePerms = Get-SmbShareAccess -Name $resolved.Share.Name -ErrorAction Stop
                foreach ($perm in $sharePerms) {
                    $results.Add([PSCustomObject]@{
                        HostName              = $serverName
                        RootFolderPath        = $drive.Root
                        SubFolderPath         = $currentNormal
                        FullNetworkPath       = $resolved.NormalUNC
                        ExtendedNetworkPath   = $resolved.ExtendedUNC
                        Owner                 = $owner
                        ACL_UserGroup         = ""
                        ACL_PermissionLevel   = ""
                        ACL_AccessType        = ""
                        InheritFrom           = ""
                        InheritanceStatus     = "N/A"
                        SharePath             = $resolved.SharePath
                        ShareName             = $resolved.ShareName
                        ShareDescription      = $resolved.ShareDescription
                        Share_UserGroup       = $perm.AccountName
                        Share_PermissionLevel = $perm.AccessRight
                        IsReparsePoint        = $false
                        LinkTarget            = ""
                    })
                }
            } catch {
                Write-Host "Share permission retrieval error: $($resolved.Share.Name)"
            }
        }

        # Enumerate child directories (include reparse points as items, but do not traverse inside them)
        try {
            $children = Get-ChildItem -LiteralPath $currentLong -Directory -Force -ErrorAction Stop
            foreach ($child in $children) {
                # Determine if child is a reparse point
                $isReparse = ($child.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0

                # Always record the child directory itself (so we don't skip anything)
                $childNormal = ConvertFrom-LongPath $child.FullName
                $childResolved = Resolve-NetworkPaths -FullPath $childNormal -ServerName $serverName -Shares $allShares

                # Record a minimal row for the child (will also be fully processed when dequeued, unless reparse)
                $results.Add([PSCustomObject]@{
                    HostName              = $serverName
                    RootFolderPath        = $drive.Root
                    SubFolderPath         = $childNormal
                    FullNetworkPath       = $childResolved.NormalUNC
                    ExtendedNetworkPath   = $childResolved.ExtendedUNC
                    Owner                 = ""
                    ACL_UserGroup         = ""
                    ACL_PermissionLevel   = ""
                    ACL_AccessType        = ""
                    InheritFrom           = ""
                    InheritanceStatus     = ""
                    SharePath             = $childResolved.SharePath
                    ShareName             = $childResolved.ShareName
                    ShareDescription      = $childResolved.ShareDescription
                    Share_UserGroup       = ""
                    Share_PermissionLevel = ""
                    IsReparsePoint        = $isReparse
                    LinkTarget            = ($child.LinkTarget -join ';')
                })

                # Only traverse into non-reparse directories to avoid cycles
                if (-not $isReparse) {
                    $queue.Enqueue($childNormal)
                }
            }
        } catch {
            Write-Host "Child enumeration error at: $currentNormal"
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
