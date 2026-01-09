
$VCServers = @(
    'Cevc01.skanska.pl',
    'Sevc10.skanska.net',
    'Sevc20.skanska.net',
    'Sevc50.skanska.net',
    'Sevc70.skanska.net',
    'USvcsa01.skanska.com'
)

$OutputFolder = 'C:\temp'
$OutputCsv    = Join-Path $OutputFolder 'VMware_HBA_Multipath_All.csv'

# Ensure output folder exists
if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

# Optional: suppress cert warnings
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false | Out-Null

# Prompt once and reuse the same credential for all vCenters
$vcCred = Get-Credential -Message "Enter credentials with read access to all listed vCenters"


function Get-PropValue {
    param(
        [Parameter(Mandatory=$true)]$Obj,
        [Parameter(Mandatory=$true)][string[]]$Candidates
    )
    if (-not $Obj) { return $null }
    $map = @{}
    foreach ($p in $Obj.PSObject.Properties) {
        $key = ($p.Name -replace '[\s_-]','').ToLower()
        $map[$key] = $p.Name
    }
    foreach ($cand in $Candidates) {
        $key = ($cand -replace '[\s_-]','').ToLower()
        if ($map.ContainsKey($key)) { return $Obj.($map[$key]) }
    }
    return $null
}


$rows = New-Object System.Collections.Generic.List[object]

foreach ($vc in $VCServers) {
    Write-Host "Connecting to vCenter $vc ..." -ForegroundColor Cyan
    try {
        $conn = Connect-VIServer -Server $vc -Credential $vcCred -ErrorAction Stop
    } catch {
        Write-Warning "Failed to connect to $vc. Skipping. Error: $($_.Exception.Message)"
        continue
    }

    try {
        # Grab ESXi hosts
        $esxiHosts = Get-VMHost | Sort-Object -Property Name

        foreach ($esxiHost in $esxiHosts) {
            Write-Host "Processing host: $($esxiHost.Name) on $vc" -ForegroundColor Green

            $hostOs = "ESXi $($esxiHost.Version) (Build $($esxiHost.Build))"
            $esxcli = Get-EsxCli -VMHost $esxiHost -V2

            # ---------------------------
            # HBA info + firmware
            # ---------------------------
            $hbas = Get-VMHostHba -VMHost $esxiHost -Type FibreChannel,iSCSI,FCoE -ErrorAction SilentlyContinue

            # Firmware per adapter via esxcli storage core adapter get/list
            $firmwareByAdapter = @{}
            $hbaInfoByAdapter  = @{}

            $adapterList = @()
            try { $adapterList = $esxcli.storage.core.adapter.list.Invoke() } catch {}

            foreach ($adapter in $adapterList) {
                $adapterName = Get-PropValue -Obj $adapter -Candidates @('Adapter','Name')
                if (-not $adapterName) { continue }
                try {
                    $getArgs = $esxcli.storage.core.adapter.get.CreateArgs()
                    $getArgs.Adapter = $adapterName
                    $adapterDetails = $esxcli.storage.core.adapter.get.Invoke($getArgs)
                    $fw = Get-PropValue -Obj $adapterDetails -Candidates @('Firmware Version','FirmwareVersion')
                    if (-not $fw) { $fw = "" }
                    $firmwareByAdapter[$adapterName] = $fw
                } catch {
                    $firmwareByAdapter[$adapterName] = ""
                }
            }

            if ($hbas) {
                foreach ($hba in $hbas) {
                    $makeModelParts = @()
                    if ($hba.Manufacturer) { $makeModelParts += $hba.Manufacturer }
                    if ($hba.Model)        { $makeModelParts += $hba.Model }
                    $makeModel = ($makeModelParts -join " ").Trim()

                    $fw = ""
                    if ($hba.Device -and $firmwareByAdapter.ContainsKey($hba.Device)) {
                        $fw = $firmwareByAdapter[$hba.Device]
                    }
                    $hbaInfoByAdapter[$hba.Device] = @{
                        MakeModel = $makeModel
                        Firmware  = $fw
                    }
                }
            }

            # ---------------------------
            # Multipath device + path lists
            # ---------------------------
            $nmpDevices = @()
            $nmpPaths   = @()
            $corePaths  = @()

            try { $nmpDevices = $esxcli.storage.nmp.device.list.Invoke() } catch {}
            try { $nmpPaths   = $esxcli.storage.nmp.path.list.Invoke() }   catch {}
            try { $corePaths  = $esxcli.storage.core.path.list.Invoke() }  catch {}

            # Device map (canonical -> SATP/PSP/etc.)
            $nmpDeviceMap = @{}
            foreach ($dev in $nmpDevices) {
                $canonical   = Get-PropValue -Obj $dev -Candidates @('Device','Canonical Name')
                if (-not $canonical) { continue }
                $satp        = Get-PropValue -Obj $dev -Candidates @('Storage Array Type Plugin','SATP')
                $psp         = Get-PropValue -Obj $dev -Candidates @('Path Selection Policy','PSP','PathSelectionPolicy')
                $pspOptions  = Get-PropValue -Obj $dev -Candidates @('Path Selection Policy Device Config','PSP Options','PathSelectionPolicyDeviceConfig')
                $dispName    = Get-PropValue -Obj $dev -Candidates @('Display Name','Device Display Name','Name')
                $opState     = Get-PropValue -Obj $dev -Candidates @('Operational State','OperationalState')

                $nmpDeviceMap[$canonical] = @{
                    SATP       = $satp
                    PSP        = $psp
                    PSPOptions = $pspOptions
                    Display    = $dispName
                    OpState    = $opState
                }
            }

            # NMP path meta keyed by runtime name
            $nmpPathMetaByRuntime = @{}
            foreach ($p in $nmpPaths) {
                $rt   = Get-PropValue -Obj $p -Candidates @('Runtime Name','RuntimeName','Name','Path')
                if (-not $rt) { continue }
                $work = Get-PropValue -Obj $p -Candidates @('Working','IsWorking','Is Working Path')
                $pref = Get-PropValue -Obj $p -Candidates @('Preferred')
                $stat = Get-PropValue -Obj $p -Candidates @('State','Path State')
                $nmpPathMetaByRuntime[$rt] = @{
                    Working   = $work
                    Preferred = $pref
                    State     = $stat
                }
            }

            # Count paths / active paths per device from corePaths
            $pathCountByDevice   = @{}
            $activeCountByDevice = @{}
            foreach ($p in $corePaths) {
                $dev = Get-PropValue -Obj $p -Candidates @('Device')
                if (-not $dev) { continue }
                if (-not $pathCountByDevice.ContainsKey($dev)) { $pathCountByDevice[$dev] = 0 }
                $pathCountByDevice[$dev]++

                $rt    = Get-PropValue -Obj $p -Candidates @('Runtime Name','RuntimeName','Name','Path')
                $state = Get-PropValue -Obj $p -Candidates @('State','Path State')
                $work  = $null
                if ($rt -and $nmpPathMetaByRuntime.ContainsKey($rt)) {
                    $work = $nmpPathMetaByRuntime[$rt].Working
                }

                $isActive = $false
                if ($state -and ($state.ToString().ToLower().Contains('active'))) { $isActive = $true }
                if ($work -and $work.ToString().ToLower() -in @('true','y','yes','working')) { $isActive = $true }

                if ($isActive) {
                    if (-not $activeCountByDevice.ContainsKey($dev)) { $activeCountByDevice[$dev] = 0 }
                    $activeCountByDevice[$dev]++
                }
            }

            # ---------------------------
            # Rows: prefer one row per PATH (full multipaths).
            # If no paths exist, emit device-only rows.
            # ---------------------------
            if ($corePaths -and $corePaths.Count -gt 0) {
                foreach ($p in $corePaths) {
                    $canonical = Get-PropValue -Obj $p -Candidates @('Device')
                    $rtName    = Get-PropValue -Obj $p -Candidates @('Runtime Name','RuntimeName','Name','Path') # vmhbaX:Cx:Tx:Lx
                    $adapter   = Get-PropValue -Obj $p -Candidates @('Adapter')                                # vmhbaX
                    $target    = Get-PropValue -Obj $p -Candidates @('Target','Target Number')
                    $lun       = Get-PropValue -Obj $p -Candidates @('LUN','Lun','Lun Id')
                    $state     = Get-PropValue -Obj $p -Candidates @('State','Path State')
                    $pref      = Get-PropValue -Obj $p -Candidates @('Preferred')
                    $transport = Get-PropValue -Obj $p -Candidates @('Transport','Transport Type')

                    # Merge NMP meta if missing
                    $work = $null
                    if ($rtName -and $nmpPathMetaByRuntime.ContainsKey($rtName)) {
                        $work = $nmpPathMetaByRuntime[$rtName].Working
                        if (-not $pref)  { $pref  = $nmpPathMetaByRuntime[$rtName].Preferred }
                        if (-not $state) { $state = $nmpPathMetaByRuntime[$rtName].State }
                    }

                    $makeModel = if ($adapter -and $hbaInfoByAdapter.ContainsKey($adapter)) { $hbaInfoByAdapter[$adapter].MakeModel } else { "" }
                    $fw        = if ($adapter -and $hbaInfoByAdapter.ContainsKey($adapter)) { $hbaInfoByAdapter[$adapter].Firmware } else { "" }

                    $satp      = ""
                    $psp       = ""
                    $pspOpt    = ""
                    $dispName  = ""
                    $opState   = ""
                    if ($canonical -and $nmpDeviceMap.ContainsKey($canonical)) {
                        $satp     = $nmpDeviceMap[$canonical].SATP
                        $psp      = $nmpDeviceMap[$canonical].PSP
                        $pspOpt   = $nmpDeviceMap[$canonical].PSPOptions
                        $dispName = $nmpDeviceMap[$canonical].Display
                        $opState  = $nmpDeviceMap[$canonical].OpState
                    }

                    $totalPaths  = if ($canonical -and $pathCountByDevice.ContainsKey($canonical))   { $pathCountByDevice[$canonical] } else { $null }
                    $activePaths = if ($canonical -and $activeCountByDevice.ContainsKey($canonical)) { $activeCountByDevice[$canonical] } else { $null }

                    $rows.Add([pscustomobject]@{
                        vCenter               = $vc
                        HostName              = $esxiHost.Name
                        HostOS                = $hostOs
                        HBA_Adapter           = $adapter
                        HBA_MakeModel         = $makeModel
                        HBA_FirmwareVersion   = $fw
                        Device_Canonical      = $canonical        # naa.*
                        Device_DisplayName    = $dispName
                        SATP                  = $satp
                        PSP                   = $psp
                        PSP_Options           = $pspOpt
                        Device_PathCount      = $totalPaths
                        Device_ActivePaths    = $activePaths
                        Device_Operational    = $opState
                        Path_RuntimeName      = $rtName           # vmhbaX:Cx:Tx:Lx
                        Path_Target           = $target
                        Path_LUN              = $lun
                        Path_State            = $state            # Active/Standby/Dead
                        Path_Preferred        = $pref
                        Path_Working          = $work
                        Path_Transport        = $transport
                    })
                }
            }
            else {
                # No corePaths; output device-only lines
                foreach ($dev in $nmpDevices) {
                    $canonical = Get-PropValue -Obj $dev -Candidates @('Device','Canonical Name')
                    if (-not $canonical) { continue }
                    $satp     = Get-PropValue -Obj $dev -Candidates @('Storage Array Type Plugin','SATP')
                    $psp      = Get-PropValue -Obj $dev -Candidates @('Path Selection Policy','PSP')
                    $pspOpt   = Get-PropValue -Obj $dev -Candidates @('Path Selection Policy Device Config','PSP Options')
                    $dispName = Get-PropValue -Obj $dev -Candidates @('Display Name','Device Display Name','Name')
                    $opState  = Get-PropValue -Obj $dev -Candidates @('Operational State','OperationalState')

                    $totalPaths  = if ($pathCountByDevice.ContainsKey($canonical))   { $pathCountByDevice[$canonical] } else { $null }
                    $activePaths = if ($activeCountByDevice.ContainsKey($canonical)) { $activeCountByDevice[$canonical] } else { $null }

                    # Adapter/HBA unknown without path; leave blank
                    $rows.Add([pscustomobject]@{
                        vCenter               = $vc
                        HostName              = $esxiHost.Name
                        HostOS                = $hostOs
                        HBA_Adapter           = ""
                        HBA_MakeModel         = ""
                        HBA_FirmwareVersion   = ""
                        Device_Canonical      = $canonical
                        Device_DisplayName    = $dispName
                        SATP                  = $satp
                        PSP                   = $psp
                        PSP_Options           = $pspOpt
                        Device_PathCount      = $totalPaths
                        Device_ActivePaths    = $activePaths
                        Device_Operational    = $opState
                        Path_RuntimeName      = ""
                        Path_Target           = ""
                        Path_LUN              = ""
                        Path_State            = ""
                        Path_Preferred        = ""
                        Path_Working          = ""
                        Path_Transport        = ""
                    })
                }
            }

        } # end foreach esxiHost
    } finally {
        Disconnect-VIServer -Server $vc -Confirm:$false | Out-Null
        Write-Host "Disconnected from $vc." -ForegroundColor DarkGray
    }
} # end foreach vCenter

# ---------------------------
# Export single CSV
# ---------------------------
$rows | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
Write-Host "All results saved to: $OutputCsv" -ForegroundColor Yellow
``
 
