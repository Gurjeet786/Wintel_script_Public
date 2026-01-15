

param(
    [int]$ExpectedPaths = 4
)

# Prepare output directory
$OutDir = 'C:\Temp'
if (-not (Test-Path $OutDir)) { New-Item -Path $OutDir -ItemType Directory -Force | Out-Null }
$timeStamp = (Get-Date).ToString('yyyyMMddHHmm')
$outCsv    = Join-Path $OutDir "Mpio_HBA_Summary_$timeStamp.csv"

Write-Host "`n=== Collecting Multipath & HBA details from local server ===" -ForegroundColor Cyan

# -------- System Info --------
$ComputerHostName = $env:COMPUTERNAME
$os   = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
$cs   = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
$model = $cs.Model
$osCap = $os.Caption
$hostOsCol = if ($model -and $osCap) { "$model / $osCap" } elseif ($osCap) { $osCap } else { $model }

# -------- HBA Info --------
$hbaMake = $null; $hbaFw = $null
try {
    $fcAttrs = Get-CimInstance -Namespace root\WMI -ClassName MSFC_FCAdapterHBAAttributes -ErrorAction Stop
    if ($fcAttrs) {
        $mk = @(); $fw = @()
        foreach ($a in $fcAttrs) {
            $mn = $a.Manufacturer; $md = $a.Model
            if ($md -and $mn) { $mk += "$md ($mn)" }
            elseif ($md)     { $mk += $md }
            elseif ($mn)     { $mk += $mn }
            if ($a.FirmwareVersion) { $fw += $a.FirmwareVersion }
            elseif ($a.DriverVersion) { $fw += $a.DriverVersion }
        }
        if ($mk.Count -gt 0) { $hbaMake = ($mk | Sort-Object -Unique) -join '; ' }
        if ($fw.Count -gt 0) { $hbaFw   = ($fw | Sort-Object -Unique) -join '; ' }
    }
} catch { }

# Fallback via PnP
if (-not $hbaMake -or -not $hbaFw) {
    $pnp = Get-CimInstance Win32_PnPEntity -ErrorAction SilentlyContinue |
           Where-Object { $_.Name -match 'Fibre|Fiber' -and $_.Name -match 'Channel' }
    if ($pnp) {
        if (-not $hbaMake) {
            $names = $pnp | Select-Object -ExpandProperty Name
            if ($names) { $hbaMake = ($names | Sort-Object -Unique) -join '; ' }
        }
        if (-not $hbaFw) {
            $drv = Get-CimInstance Win32_PnPSignedDriver -ErrorAction SilentlyContinue |
                   Where-Object { $_.DeviceName -in ($pnp | Select-Object -ExpandProperty Name) }
            $ver = $drv | Select-Object -ExpandProperty DriverVersion
            if ($ver) { $hbaFw = ($ver | Sort-Object -Unique) -join '; ' }
        }
    }
}

# -------- Multipath Detection --------
$mpclaimText = $null
try { $mpclaimText = (& mpclaim.exe -s -d 2>$null) } catch { }
$dsmNames = @()
if ($mpclaimText) {
    foreach ($line in $mpclaimText) {
        if ($line -match 'DSM\s+Name\s*:\s*(.+)$') { $dsmNames += $Matches[1].Trim() }
    }
}
$isPowerPath = ($dsmNames -match 'PowerPath')

# Collect path info
$pathRecords = @()
if ($isPowerPath) {
    $powermt = $null
    if (Get-Command powermt.exe -ErrorAction SilentlyContinue) {
        try { $powermt = (& powermt.exe display dev=all 2>$null) } catch { }
    }
    if ($powermt) {
        $curInst = $null; $act=0; $std=0; $fai=0; $tot=0
        foreach ($line in $powermt) {
            $l = $line.Trim()
            if ($l -match '^(Pseudo\s+name|Logical\s+device|Host\s+Device)\s*[:=]\s*(.+)$') {
                if ($curInst) {
                    $pathRecords += [pscustomobject]@{ Instance=$curInst; PathCount=$tot; Active=$act; Standby=$std; Failed=$fai }
                }
                $curInst = $Matches[2].Trim(); $act=0; $std=0; $fai=0; $tot=0
                continue
            }
            if ($l -match '(?i)\b(active|standby|alive|dead|failed)\b') {
                $tot++
                if ($l -match '(?i)\bactive\b' -or $l -match '(?i)\balive\b') { $act++ }
                elseif ($l -match '(?i)\bstandby\b') { $std++ }
                elseif ($l -match '(?i)\bdead\b' -or $l -match '(?i)\bfailed\b') { $fai++ }
            }
        }
        if ($curInst) {
            $pathRecords += [pscustomobject]@{ Instance=$curInst; PathCount=$tot; Active=$act; Standby=$std; Failed=$fai }
        }
    }
} elseif (Get-Command Get-MPIOPath -ErrorAction SilentlyContinue) {
    $mpioObjs = Get-MPIOPath
    $groups = $mpioObjs | Group-Object InstanceName
    foreach ($g in $groups) {
        $paths = $g.Group
        $act = ($paths | Where-Object {$_.State -match 'Active'}).Count
        $std = ($paths | Where-Object {$_.State -match 'Standby'}).Count
        $fai = ($paths | Where-Object {$_.State -match 'Failed'}).Count
        $tot = $paths.Count
        $pathRecords += [pscustomobject]@{ Instance=$g.Name; PathCount=$tot; Active=$act; Standby=$std; Failed=$fai }
    }
}

# Aggregate
$pathsArr = $pathRecords | Select-Object -ExpandProperty PathCount
$activeArr = $pathRecords | Select-Object -ExpandProperty Active
$failedSum = ($pathRecords | Measure-Object -Property Failed -Sum).Sum
$modePaths = if ($pathsArr) { ($pathsArr | Group-Object | Sort-Object Count -Descending | Select-Object -First 1).Name } else { $null }
$modeActive = if ($activeArr) { ($activeArr | Group-Object | Sort-Object Count -Descending | Select-Object -First 1).Name } else { $null }
$mpStatus = if ($modePaths) { "$modePaths paths, $modeActive active" } else { "N/A" }

# Comments
$comments = @()
if ($modePaths -and ($modePaths -ne $ExpectedPaths)) { $comments += "Expected $ExpectedPaths, found $modePaths" }
if ($failedSum -and $failedSum -gt 0) { $comments += "$failedSum failed path(s)" }
$commentText = ($comments -join '; ')

# Final object
$result = [pscustomobject]@{
    'Host name'              = $ComputerHostName
    'HOST OS (Esxi Version)' = $hostOsCol
    'Host HBA Make'          = $hbaMake
    'HBA firmware version'   = $hbaFw
    'Multipath status'       = $mpStatus
    'Comments'               = $commentText
}

# Output
$result | Format-Table -AutoSize
$result | Export-Csv -Path $outCsv -NoTypeInformation -Encoding UTF8
Write-Host "`nCSV saved to: $outCsv" -ForegroundColor Green
