
# --- 1. Prereqs & connection -------------------------------------------------

if (-not (Get-Module -ListAvailable -Name VMware.PowerCLI)) {
    Write-Error "VMware.PowerCLI module not found. Install-Module VMware.PowerCLI first."
    return
}

Import-Module VMware.PowerCLI -ErrorAction Stop
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false | Out-Null

$vCenter = Read-Host "Enter vCenter Server FQDN or IP"
$cred    = Get-Credential -Message "Enter vCenter credentials"

Connect-VIServer -Server $vCenter -Credential $cred | Out-Null

# --- 2. Output path ----------------------------------------------------------

$csvPath = "C:\Temp\ESXi-Multipath-Detailed.csv"
if (-not (Test-Path "C:\Temp")) {
    New-Item -Path "C:\Temp" -ItemType Directory -Force | Out-Null
}

Write-Host ""
Write-Host "===== Loading inventory from $vCenter ====="

# --- 3. Inventory (no filters first, just count) -----------------------------

$allHosts      = Get-VMHost
$allDatastores = Get-Datastore          # do not filter yet
$allVMs        = Get-VM

Write-Host "Hosts:       $($allHosts.Count)"
Write-Host "Datastores:  $($allDatastores.Count)"
Write-Host "VMs:         $($allVMs.Count)"
Write-Host ""

if ($allHosts.Count -eq 0) {
    Write-Warning "No ESXi hosts found in inventory. Exiting."
    Disconnect-VIServer -Server $vCenter -Confirm:$false | Out-Null
    return
}

# Keep only block (VMFS) datastores for LUN/path logic, but
# we already printed the unfiltered count above.
$vmfsDatastores = $allDatastores | Where-Object { $_.Type -eq "VMFS" }

Write-Host "VMFS datastores (block): $($vmfsDatastores.Count)"
Write-Host ""

# --- 4. Build lookup tables --------------------------------------------------

# NAA (canonical name) -> datastore names
$naaToDatastore = @{}
foreach ($ds in $vmfsDatastores) {
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

# Datastore -> list of VM + OS
$dsToVMInfo = @{}
foreach ($vm in $allVMs) {
    foreach ($ds in $vm.Datastore) {
        if (-not $dsToVMInfo.ContainsKey($ds.Name)) {
            $dsToVMInfo[$ds.Name] = @()
        }
        $dsToVMInfo[$ds.Name] += [PSCustomObject]@{
            VMName  = $vm.Name
            GuestOS = $vm.Guest.OSFullName
        }
    }
}

# --- 5. Main multipath collection -------------------------------------------

$results   = @()
$hostCount = $allHosts.Count
$idx       = 0

foreach ($esxiHost in $allHosts) {

    $idx++
    Write-Progress -Activity "Collecting multipath details" `
                   -Status "Host $idx of $hostCount : $($esxiHost.Name)" `
                   -PercentComplete (($idx / $hostCount) * 100)

    Write-Host "Host: $($esxiHost.Name)"

    $hostView    = $esxiHost.ExtensionData
    $hw          = $hostView.Hardware.SystemInfo
    $esxiVersion = $esxiHost.Version
    $esxiBuild   = $esxiHost.Build

    # IMPORTANT: no -Type filter here so we cannot miss HBAs
    $hbas = Get-VMHostHba -VMHost $esxiHost -ErrorAction SilentlyContinue
    Write-Host "  HBAs found: $($hbas.Count)"

    if (-not $hbas -or $hbas.Count -eq 0) { continue }

    foreach ($hba in $hbas) {

        $hbaExt    = $hba.ExtensionData
        $firmware  = $null
        $driver    = $null
        $model     = $hba.Model
        $wwpn      = $null

        if ($hbaExt) {
            if ($hbaExt.Driver)          { $driver   = $hbaExt.Driver }
            if ($hbaExt.FirmwareVersion) { $firmware = $hbaExt.FirmwareVersion }
            if ($hbaExt.PortWorldWideName) {
                $wwpn = ("{0:x}" -f $hbaExt.PortWorldWideName) -split '([a-f0-9]{2})' |
                        Where-Object { $_ } | ForEach-Object { $_ } -join ":"
            }
        }

        $luns = Get-ScsiLun -Hba $hba -ErrorAction SilentlyContinue
        Write-Host "    HBA $($hba.Device) LUNs: $($luns.Count)"

        if (-not $luns -or $luns.Count -eq 0) { continue }

        foreach ($lun in $luns) {

            $canonicalName = $lun.CanonicalName
            $lunId         = (($lun.RuntimeName -split "L")[1] -as [int])

            # Map LUN -> datastore names
            $dsNames = @()
            if ($canonicalName -and $naaToDatastore.ContainsKey($canonicalName)) {
                $dsNames = $naaToDatastore[$canonicalName] | Sort-Object -Unique
            }

            # Aggregate VM and OS info for those datastores
            $vmNames  = @()
            $vmOSList = @()
            foreach ($dsName in $dsNames) {
                if ($dsToVMInfo.ContainsKey($dsName)) {
                    foreach ($vmInfo in $dsToVMInfo[$dsName]) {
                        $vmNames  += $vmInfo.VMName
                        if ($vmInfo.GuestOS) {
                            $vmOSList += "$($vmInfo.VMName): $($vmInfo.GuestOS)"
                        }
                    }
                }
            }
            $vmNames  = $vmNames  | Sort-Object -Unique
            $vmOSList = $vmOSList | Sort-Object -Unique

            # Paths for this LUN
            $paths = $lun | Get-ScsiLunPath -ErrorAction SilentlyContinue
            Write-Host "      LUN $canonicalName paths: $($paths.Count)"

            if (-not $paths -or $paths.Count -eq 0) { continue }

            $totalPaths  = $paths.Count
            $activePaths = ($paths | Where-Object { $_.State -eq "Active" }).Count
            $mpSummary   = if ($totalPaths -gt 0) {
                               "$totalPaths paths, $activePaths active"
                           } else {
                               "No paths"
                           }

            foreach ($path in $paths) {

                $results += [PSCustomObject]@{
                    vCenter              = $vCenter
                    Cluster              = $esxiHost.Parent.Name
                    ESXiHost             = $esxiHost.Name
                    ESXiVersion          = $esxiVersion
                    ESXiBuild            = $esxiBuild
                    HostVendor           = $hw.Vendor
                    HostModel            = $hw.Model

                    # multiple Windows physical servers / VMs using that storage
                    HostOS_VMGuest       = ($vmOSList -join "; ")

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
                    PathState            = $path.State
                    PathTarget           = $path.SanID
                    PathLun              = $path.ExtensionData.Lun
                    PathStatus           = $path.ExtensionData.PathState
                }
            }
        }
    }
}

# --- 6. Export or warn -------------------------------------------------------

Write-Host ""
Write-Host "Total path records collected: $($results.Count)"

if ($results.Count -gt 0) {
    $results |
        Sort-Object Cluster, ESXiHost, CanonicalName, PathName |
        Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Detailed multipath report written to $csvPath"
} else {
    Write-Warning "No multipath data collected. Check HBA types, storage type (VMFS vs NFS), and permissions."
}

Disconnect-VIServer -Server $vCenter -Confirm:$false | Out-Null
