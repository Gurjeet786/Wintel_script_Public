$vCenters = @(
    "Cevc01.skanska.pl",
    "Sevc10.skanska.net",
    "Sevc20.skanska.net",
    "Sevc50.skanska.net",
    "Sevc70.skanska.net",
    "USvcsa01.skanska.com"
)

# Allow multiple connections [web:38]
Set-PowerCLIConfiguration -DefaultVIServerMode Multiple -Confirm:$false | Out-Null

# Connect to each vCenter, prompting for creds every time [web:36][web:44]
foreach ($vc in $vCenters) {
    try {
        Write-Host "Connecting to $vc..." -ForegroundColor Cyan
        $creds = Get-Credential -Message "Enter credentials for $vc"
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

# All hosts from all connected vCenters [web:17][web:23]
$vmhosts = Get-VMHost  

$report = @()

foreach ($vmhost in $vmhosts) {
    $hostView = Get-View $vmhost
    $hostOS   = "{0} {1} (Build {2})" -f $hostView.Config.Product.FullName,
                                       $hostView.Config.Product.Version,
                                       $hostView.Config.Product.Build

    # vCenter of this host – substring after '@' in Uid [web:24]
    $vcName = ($vmhost.Uid -split '@')[-1]

    $esxcli = Get-EsxCli -VMHost $vmhost -V2

    # FC & SAS HBA info [web:31][web:43]
    $fcList  = $esxcli.storage.san.fc.list.Invoke()
    $sasList = $esxcli.storage.san.sas.list.Invoke()

    $hbaList = @()
    if ($fcList)  { $hbaList += $fcList }
    if ($sasList) { $hbaList += $sasList }

    # Multipath info (per device) [web:9][web:34][web:40]
    $nmpDevices = $esxcli.storage.nmp.device.list.Invoke()

    if (-not $hbaList) {
        # Still output a line so host is visible even without HBA
        $report += [pscustomobject]@{
            vCenter         = $vcName
            HostName        = $vmhost.Name
            HostOS          = $hostOS
            HbaName         = $null
            HbaModel        = "No FC/SAS HBA found"
            HbaVendor       = $null
            HbaDriver       = $null
            HbaFirmware     = $null
            MultipathDevice = $null
            MultipathSATP   = $null
            MultipathPSP    = $null
        }
        continue
    }

    foreach ($hba in $hbaList) {
        # Normalise ESXCLI properties into discrete columns [web:31][web:43]
        $hbaName     = $null
        $hbaModel    = $null
        $hbaVendor   = $null
        $hbaDriver   = $null
        $hbaFirmware = $null

        if ($hba.PSObject.Properties.Name -contains 'Adapter') {
            $hbaName = $hba.Adapter
        }
        if ($hba.PSObject.Properties.Name -contains 'Model') {
            $hbaModel = $hba.Model
        } elseif ($hba.PSObject.Properties.Name -contains 'Description') {
            $hbaModel = $hba.Description
        }
        if ($hba.PSObject.Properties.Name -contains 'Vendor') {
            $hbaVendor = $hba.Vendor
        }
        if ($hba.PSObject.Properties.Name -contains 'DriverName') {
            $hbaDriver = $hba.DriverName
        } elseif ($hba.PSObject.Properties.Name -contains 'Driver') {
            $hbaDriver = $hba.Driver
        }
        if ($hba.PSObject.Properties.Name -contains 'FirmwareVersion') {
            $hbaFirmware = $hba.FirmwareVersion
        } elseif ($hba.PSObject.Properties.Name -contains 'Firmware') {
            $hbaFirmware = $hba.Firmware
        }

        # For each HBA, create one row per attached device + SATP/PSP [web:34][web:40]
        foreach ($dev in $nmpDevices) {
            $deviceId = $null
            $satp     = $null
            $psp      = $null

            if ($dev.PSObject.Properties.Name -contains 'Device') {
                $deviceId = $dev.Device
            }
            if ($dev.PSObject.Properties.Name -contains 'StorageArrayType') {
                $satp = $dev.StorageArrayType
            } elseif ($dev.PSObject.Properties.Name -contains 'StorageArrayTypePlugin') {
                $satp = $dev.StorageArrayTypePlugin
            }
            if ($dev.PSObject.Properties.Name -contains 'PathSelectionPolicy') {
                $psp = $dev.PathSelectionPolicy
            }

            # Only output rows that have a device and multipath policy
            if ($deviceId -and $satp -and $psp) {
                $report += [pscustomobject]@{
                    vCenter         = $vcName
                    HostName        = $vmhost.Name
                    HostOS          = $hostOS
                    HbaName         = $hbaName
                    HbaModel        = $hbaModel
                    HbaVendor       = $hbaVendor
                    HbaDriver       = $hbaDriver
                    HbaFirmware     = $hbaFirmware
                    MultipathDevice = $deviceId
                    MultipathSATP   = $satp
                    MultipathPSP    = $psp
                }
            }
        }
    }
}

$report | Export-Csv -Path $outFile -NoTypeInformation -UseCulture

Write-Host "Report written to $outFile"
