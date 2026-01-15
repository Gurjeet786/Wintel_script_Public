
if (-not (Get-Module -ListAvailable -Name VMware.PowerCLI)) {
    Write-Error "VMware.PowerCLI module not found. Install-Module VMware.PowerCLI first."
    return
}

Import-Module VMware.PowerCLI -ErrorAction Stop
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false | Out-Null

$vCenter = Read-Host "Enter vCenter Server FQDN or IP"
$cred    = Get-Credential -Message "Enter vCenter credentials"

Connect-VIServer -Server $vCenter -Credential $cred | Out-Null

# Output path
$csvPath = "C:\Temp\ESXi-Multipath-Detailed.csv"
if (-not (Test-Path "C:\Temp")) {
    New-Item -Path "C:\Temp" -ItemType Directory -Force | Out-Null
}

Write-Host "Loading inventory from $vCenter ..."

$allHosts      = Get-VMHost
$allDatastores = Get-Datastore | Where-Object { $_.Type -eq "VMFS" }   # block storage only [web:2][page:1]
$allVMs        = Get-VM

# Map NAA (canonical name) -> datastore names
$naaToDatastore = @{}
foreach ($ds in $allDatastores) {
    $vmfs = $ds.ExtensionData.Info.Vmfs
    if ($vmfs -and $vmfs.Extent) {
        foreach ($ext in $vmfs.Extent) {
            $naa = $ext.DiskName
            if ([string]::IsNullOrWhiteSpace($naa)) { continue }
            if (-not $naaToDatastore.ContainsKey($naa)) {
                $naaToDatastore[$naa] = @()
            }
            $naaToDatastore[$naa] += $ds.Name
        }
    }
}

# Map datastore -> VMs & their OS
$dsToVMInfo = @{}
foreach ($vm in $allVMs) {
    foreach ($ds in $vm.Datastore) {
        if (-not $dsToVMInfo.ContainsKey($ds.Name)) {
            $dsToVMInfo[$ds.Name] = @()
        }
        $guestOS = $vm.Guest.OSFullName
        $dsToVMInfo[$ds.Name] += [PSCustomObject]@{
            VMName  = $vm.Name
            GuestOS = $guestOS
        }
    }
}

$results = @()
$hostCount = $allHosts.Count
$idx = 0

foreach ($esxiHost in $allHosts) {
    $idx++
    Write-Progress -Activity "Collecting multipath details" `
                   -Status "Host $idx of $hostCount : $($esxiHost.Name)" `
                   -PercentComplete (($idx / $hostCount) * 100)

    $hostView    = $esxiHost.ExtensionData
    $hw          = $hostView.Hardware.SystemInfo
    $esxiVersion = $esxiHost.Version
    $esxiBuild   = $esxiHost.Build

    $hbas = Get-VMHostHba -VMHost $esxiHost -Type FibreChannel, iSCSI, FCoE

    foreach ($hba in $hbas) {

        # Try to resolve firmware and WWPN from ExtensionData where available [page:1]
        $hbaExt = $hba.ExtensionData
        $firmware = $null
        $driver   = $null
        $model    = $hba.Model
        $wwpn     = $null

        if ($hbaExt -and $hbaExt.Driver) { $driver = $hbaExt.Driver }
        if ($hbaExt -and $hbaExt.FirmwareVersion) { $firmware = $hbaExt.FirmwareVersion }
        if ($hbaExt -and $hbaExt.PortWorldWideName) {
            $wwpn = ("{0:x}" -f $hbaExt.PortWorldWideName) -split '([a-f0-9]{2})' | Where-Object { $_ } | ForEach-Object { $_ } -join ":"
        }

        $luns = Get-ScsiLun -Hba $hba -ErrorAction SilentlyContinue
        if (-not $luns) { continue }

        foreach ($lun in $luns) {

            $canonicalName = $lun.CanonicalName
            $lunId         = (($lun.RuntimeName -split "L")[1] -as [int])

            # Datastores for this LUN
            $dsNames = @()
            if ($canonicalName -and $naaToDatastore.ContainsKey($canonicalName)) {
                $dsNames = $naaToDatastore[$canonicalName] | Sort-Object -Unique
            }

            # Aggregate VM & OS info
            $vmNames  = @()
            $vmOSList = @()
            foreach ($dsName in $dsNames) {
                if ($dsToVMInfo.ContainsKey($dsName)) {
                    foreach ($vmInfo in $dsToVMInfo[$dsName]) {
                        $vmNames  += $vmInfo.VMName
                        if ($vmInfo.GuestOS) { $vmOSList += "$($vmInfo.VMName): $($vmInfo.GuestOS)" }
                    }
                }
            }
            $vmNames  = $vmNames  | Sort-Object -Unique
            $vmOSList = $vmOSList | Sort-Object -Unique

            # Get all paths for this LUN
            $paths = $lun | Get-ScsiLunPath -ErrorAction SilentlyContinue
            if (-not $paths) { continue }

            $totalPaths  = $paths.Count
            $activePaths = ($paths | Where-Object { $_.State -eq "Active" }).Count   # [web:2][page:1]
            $mpSummary   = if ($totalPaths -gt 0) { "$totalPaths paths, $activePaths active" } else { "No paths" }

            foreach ($path in $paths) {

                $obj = [PSCustomObject]@{
                    vCenter              = $vCenter
                    Cluster              = $esxiHost.Parent.Name
                    ESXiHost             = $esxiHost.Name
                    ESXiVersion          = $esxiVersion
                    ESXiBuild            = $esxiBuild
                    HostVendor           = $hw.Vendor
                    HostModel            = $hw.Model

                    HostOS_VMGuest       = ($vmOSList -join "; ")   # multiple Windows servers strings

                    HBADevice            = $hba.Device
                    HBAType              = $hba.Type
                    HBAModel             = $model
                    HBADriver            = $driver
                    HBAFirmwareVersion   = $firmware
                    HBA_WWPN             = $wwpn

                    CanonicalName        = $canonicalName
                    LUNID                = $lunId
                    MultipathPolicy      = $lun.MultipathPolicy
                    TotalPaths           = $totalPaths
                    ActivePaths          = $activePaths
                    MultipathStatusText  = $mpSummary

                    Datastores           = ($dsNames -join "; ")
                    VMs                  = ($vmNames -join "; ")

                    PathName             = $path.Name
                    PathRuntimeName      = $path.RuntimeName
                    PathState            = $path.State                # Active/Standby/Dead
                    PathTarget           = $path.SanID
                    PathLun              = $path.ExtensionData.Lun
                    PathStatus           = $path.ExtensionData.PathState
                }

                $results += $obj
            }
        }
    }
}

if ($results.Count -gt 0) {
    $results |
        Sort-Object Cluster, ESXiHost, CanonicalName, PathName |
        Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    Write-Host "Detailed multipath report written to $csvPath"
} else {
    Write-Warning "No multipath data collected. Check permissions or filters."
}

Disconnect-VIServer -Server $vCenter -Confirm:$false | Out-Null
