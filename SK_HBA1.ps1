

# --- List of vCenters ---
$vCenters = @(
    "Cevc01.skanska.pl",
    "Sevc10.skanska.net",
    "Sevc20.skanska.net",
    "Sevc50.skanska.net",
    "Sevc70.skanska.net",
    "USvcsa01.skanska.com"
)

# Ask for a single credential used for all vCenters
$creds = Get-Credential

# Enable multiple VIServer connections if not already configured [web:26]
Set-PowerCLIConfiguration -DefaultVIServerMode Multiple -Confirm:$false | Out-Null

# Connect to all vCenters [web:17][web:23]
foreach ($vc in $vCenters) {
    try {
        Write-Host "Connecting to $vc..." -ForegroundColor Cyan
        Connect-VIServer -Server $vc -Credential $creds -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Warning "Failed to connect to $vc : $($_.Exception.Message)"
    }
}

# Ensure output folder exists
$outFolder = "C:\temp"
if (-not (Test-Path $outFolder)) {
    New-Item -Path $outFolder -ItemType Directory | Out-Null
}
$outFile = Join-Path $outFolder "ESXi_HBA_Multipath_Info.csv"

$vmhosts = Get-VMHost   # This will span all connected vCenters [web:17][web:23]

$report = @()

foreach ($vmhost in $vmhosts) {
    $hostView = Get-View $vmhost
    $hostOS   = "{0} {1} (Build {2})" -f $hostView.Config.Product.FullName,
                                       $hostView.Config.Product.Version,
                                       $hostView.Config.Product.Build

    $esxcli = Get-EsxCli -VMHost $vmhost -V2

    # HBA info from FC/SAS lists [web:7]
    $fcList  = $esxcli.storage.san.fc.list.Invoke()
    $sasList = $esxcli.storage.san.sas.list.Invoke()

    $hbaList = @()
    if ($fcList)  { $hbaList += $fcList }
    if ($sasList) { $hbaList += $sasList }

    # Multipath info (SATP/PSP) [web:15]
    $nmpDevices = $esxcli.storage.nmp.device.list.Invoke()

    if (-not $hbaList) {
        $report += [pscustomobject]@{
            vCenter           = $vmhost.Uid.Split("@")[1]   # origin vCenter [web:23]
            HostName          = $vmhost.Name
            HostOS            = $hostOS
            HbaDevice         = ""
            HbaMake           = "No FC/SAS HBA found"
            HbaFirmware       = ""
            MultipathDevices  = ""
            MultipathSATP_PSP = ""
        }
        continue
    }

    foreach ($hba in $hbaList) {
        $hbaDevice   = $hba.Adapter | Out-String
        $hbaModel    = $null
        $hbaFirmware = $null

        if ($hba.PSObject.Properties.Name -contains 'Model') {
            $hbaModel = $hba.Model
        } elseif ($hba.PSObject.Properties.Name -contains 'Description') {
            $hbaModel = $hba.Description
        }

        if ($hba.PSObject.Properties.Name -contains 'FirmwareVersion') {
            $hbaFirmware = $hba.FirmwareVersion
        } elseif ($hba.PSObject.Properties.Name -contains 'Firmware') {
            $hbaFirmware = $hba.Firmware
        }

        $mpEntries = @()
        foreach ($dev in $nmpDevices) {
            $satp = $null
            $psp  = $null

            if ($dev.PSObject.Properties.Name -contains 'StorageArrayType') {
                $satp = $dev.StorageArrayType
            } elseif ($dev.PSObject.Properties.Name -contains 'StorageArrayTypePlugin') {
                $satp = $dev.StorageArrayTypePlugin
            }

            if ($dev.PSObject.Properties.Name -contains 'PathSelectionPolicy') {
                $psp = $dev.PathSelectionPolicy
            }

            $deviceId = $null
            if ($dev.PSObject.Properties.Name -contains 'Device') {
                $deviceId = $dev.Device
            }

            if ($deviceId -and $satp -and $psp) {
                $mpEntries += "{0}: {1}/{2}" -f $deviceId, $satp, $psp
            }
        }

        $mpDevices  = ($mpEntries | ForEach-Object { ($_ -split ':')[0] } | Select-Object -Unique) -join '; '
        $mpSatpPsp  = ($mpEntries | Select-Object -Unique) -join '; '

        $report += [pscustomobject]@{
            vCenter           = $vmhost.Uid.Split("@")[1]
            HostName          = $vmhost.Name
            HostOS            = $hostOS
            HbaDevice         = $hbaDevice.Trim()
            HbaMake           = $hbaModel
            HbaFirmware       = $hbaFirmware
            MultipathDevices  = $mpDevices
            MultipathSATP_PSP = $mpSatpPsp
        }
    }
}

$report | Export-Csv -Path $outFile -NoTypeInformation -UseCulture

Write-Host "Report written to $outFile"
