
$vCenters = @(List of vcenter)


Set-PowerCLIConfiguration -DefaultVIServerMode Multiple -Confirm:$false | Out-Null

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

$outFolder = "C:\temp"
if (-not (Test-Path $outFolder)) {
    New-Item -Path $outFolder -ItemType Directory | Out-Null
}
$outFile = Join-Path $outFolder "ESXi_HBA_Multipath_Full.csv"

$vmhosts = Get-VMHost
$report  = @()

foreach ($vmhost in $vmhosts) {
    $hostView = Get-View $vmhost
    $hostOS   = "{0} {1} (Build {2})" -f $hostView.Config.Product.FullName,
                                       $hostView.Config.Product.Version,
                                       $hostView.Config.Product.Build

    $vcName = ($vmhost.Uid -split '@')[-1]

    $esxcli = Get-EsxCli -VMHost $vmhost -V2

    # Core adapter list – accurate HBA info. [web:31][web:46]
    $coreAdapters = $esxcli.storage.core.adapter.list.Invoke()

    # FC + SAS lists – firmware version. [web:46][web:7]
    $fcList  = $esxcli.storage.san.fc.list.Invoke()
    $sasList = $esxcli.storage.san.sas.list.Invoke()

    $fwByAdapter = @{}

    foreach ($fc in ($fcList | Where-Object { $_.PSObject.Properties.Name -contains 'Adapter' })) {
        if ($fc.Adapter -and $fc.PSObject.Properties.Name -contains 'FirmwareVersion') {
            $fwByAdapter[$fc.Adapter] = $fc.FirmwareVersion
        }
    }

    foreach ($sas in ($sasList | Where-Object { $_.PSObject.Properties.Name -contains 'Adapter' })) {
        if ($sas.Adapter -and $sas.PSObject.Properties.Name -contains 'FirmwareVersion') {
            $fwByAdapter[$sas.Adapter] = $sas.FirmwareVersion
        }
    }

    # NMP device list – SATP/PSP and some LUN metadata. [web:47][web:15]
    $nmpDevices = $esxcli.storage.nmp.device.list.Invoke()

    # Core device list – display name / runtime name. [web:48]
    $coreDevices = $esxcli.storage.core.device.list.Invoke()

    # Build lookup for LUN details per canonical name.
    $lunInfoByDevice = @{}
    foreach ($dev in $coreDevices) {
        $devName = $dev.Device
        if (-not $devName) { continue }

        $displayName = $null
        $arrayVendor = $null
        $arrayModel  = $null

        if ($dev.PSObject.Properties.Name -contains 'DisplayName') {
            $displayName = $dev.DisplayName
        } elseif ($dev.PSObject.Properties.Name -contains 'DeviceDisplayName') {
            $displayName = $dev.DeviceDisplayName
        } elseif ($dev.PSObject.Properties.Name -contains 'CanonicalName') {
            $displayName = $dev.CanonicalName
        }

        if ($dev.PSObject.Properties.Name -contains 'Vendor') {
            $arrayVendor = $dev.Vendor
        }
        if ($dev.PSObject.Properties.Name -contains 'Model') {
            $arrayModel = $dev.Model
        }

        $lunInfoByDevice[$devName] = [ordered]@{
            DisplayName = $displayName
            Vendor      = $arrayVendor
            Model       = $arrayModel
        }
    }

    # Core path list – path‑level state. [web:15][web:48]
    $corePaths = $esxcli.storage.core.path.list.Invoke()

    # Build per‑device multipath status. [web:15][web:32]
    $mpStatusByDevice = @{}
    foreach ($p in $corePaths) {
        $dev  = $p.Device
        if (-not $dev) { continue }

        if (-not $mpStatusByDevice.ContainsKey($dev)) {
            $mpStatusByDevice[$dev] = [ordered]@{
                TotalPaths  = 0
                ActivePaths = 0
                States      = @{}
            }
        }

        $info = $mpStatusByDevice[$dev]
        $info.TotalPaths++

        $state = $null
        if ($p.PSObject.Properties.Name -contains 'State') {
            $state = $p.State
        } elseif ($p.PSObject.Properties.Name -contains 'PathState') {
            $state = $p.PathState
        }
        if ([string]::IsNullOrWhiteSpace($state)) {
            $state = "Unknown"
        }

        if (-not $info.States.ContainsKey($state)) {
            $info.States[$state] = 0
        }
        $info.States[$state]++

        if ($p.PSObject.Properties.Name -contains 'IsActive' -and $p.IsActive) {
            $info.ActivePaths++
        } elseif ($state -match 'active') {
            $info.ActivePaths++
        }

        $mpStatusByDevice[$dev] = $info
    }

    if (-not $coreAdapters) {
        $report += [pscustomobject]@{
            vCenter          = $vcName
            HostName         = $vmhost.Name
            HostOS           = $hostOS
            HbaName          = $null
            HbaModel         = "No storage adapters found"
            HbaVendor        = $null
            HbaDriver        = $null
            HbaFirmware      = $null
            LunCanonicalName = $null
            LunDisplayName   = $null
            ArrayVendor      = $null
            ArrayModel       = $null
            MultipathSATP    = $null
            MultipathPSP     = $null
            MultipathStatus  = $null
        }
        continue
    }

    foreach ($adapter in $coreAdapters) {
        $hbaName   = $adapter.Name
        $hbaModel  = $adapter.Model
        $hbaVendor = $adapter.Vendor
        $hbaDriver = $adapter.Driver
        $hbaFW     = $null

        if ($hbaName -and $fwByAdapter.ContainsKey($hbaName)) {
            $hbaFW = $fwByAdapter[$hbaName]
        }

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

            if (-not ($deviceId -and $satp -and $psp)) { continue }

            $lunDisplay = $null
            $arrayVend  = $null
            $arrayModel = $null

            if ($lunInfoByDevice.ContainsKey($deviceId)) {
                $info       = $lunInfoByDevice[$deviceId]
                $lunDisplay = $info.DisplayName
                $arrayVend  = $info.Vendor
                $arrayModel = $info.Model
            }

            $mpStatus = $null
            if ($mpStatusByDevice.ContainsKey($deviceId)) {
                $info       = $mpStatusByDevice[$deviceId]
                $stateParts = @()
                foreach ($k in $info.States.Keys) {
                    $stateParts += "{0}x {1}" -f $info.States[$k], $k
                }
                $statesStr  = $stateParts -join ', '
                $mpStatus   = "{0} paths, {1} active, {2}" -f $info.TotalPaths, $info.ActivePaths, $statesStr
            }

            $report += [pscustomobject]@{
                vCenter          = $vcName
                HostName         = $vmhost.Name
                HostOS           = $hostOS
                HbaName          = $hbaName
                HbaModel         = $hbaModel
                HbaVendor        = $hbaVendor
                HbaDriver        = $hbaDriver
                HbaFirmware      = $hbaFW
                LunCanonicalName = $deviceId
                LunDisplayName   = $lunDisplay
                ArrayVendor      = $arrayVend
                ArrayModel       = $arrayModel
                MultipathSATP    = $satp
                MultipathPSP     = $psp
                MultipathStatus  = $mpStatus
            }
        }
    }
}

$report | Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8 -Delimiter ','

Write-Host "Report written to $outFile"
