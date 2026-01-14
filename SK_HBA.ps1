

$VCServers = @(List of vcenter)


$OutputFolder = 'C:\temp'
$OutputCsv    = Join-Path $OutputFolder 'VMware_HBA_Multipath_All.csv'
$LogFile      = Join-Path $OutputFolder 'VMware_HBA_Multipath.log'

# Filters (leave empty/blank to disable)
$ClusterFilterNames    = @()       # e.g., @('Prod-Cluster-A','Prod-Cluster-B')
$DatacenterFilterNames = @()       # e.g., @('DC-East','DC-West')
$HostNameRegex         = '.*'      # e.g., '^.*prod.*$' ; default matches all

# ---------------------------
# Prep
# ---------------------------
if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

# Reset log
"[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Starting collection" | Out-File -FilePath $LogFile -Encoding UTF8

# Suppress certificate prompts (optional)
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false | Out-Null

# ---------------------------
# Helpers
# ---------------------------
function Write-Log {
    param([string]$Message)
    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message" | Out-File -FilePath $LogFile -Append -Encoding UTF8
}

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

function Try-Esxcli {
    param([scriptblock]$Block)
    try { & $Block } catch { Write-Log "ESXCLI call failed: $($_.Exception.Message)"; @() }
}

# ---------------------------
# Data structures
# ---------------------------
$rows = New-Object System.Collections.Generic.List[object]

# Summary per vCenter
$summaryByVC = @{} # vc -> @{HostsProcessed=0; HostsWithDeadPaths=0; TotalPaths=0; DeadPaths=0; PSPCounts=@{}; SATPCounts=@{}; HostsDeadPathsList=@() }

# ---------------------------
# Main loop per vCenter
# ---------------------------
foreach ($vc in $VCServers) {
    Write-Host "Connecting to vCenter $vc ..." -ForegroundColor Cyan
    Write-Log   "Connecting to vCenter $vc ..."

    $vcCred = Get-Credential -Message "Enter credentials for $vc"
    try {
        Connect-VIServer -Server $vc -Credential $vcCred -ErrorAction Stop | Out-Null
        Write-Log "Connected to $vc"
    } catch {
        Write-Warning "Failed to connect to $vc: $($_.Exception.Message)"
        Write-Log     "Failed to connect to $vc: $($_.Exception.Message)"
        # Initialize summary and emit info row
        $summaryByVC[$vc] = @{
            HostsProcessed     = 0
            HostsWithDeadPaths = 0
            TotalPaths         = 0
            DeadPaths          = 0
            PSPCounts          = @{}
            SATPCounts         = @{}
            HostsDeadPathsList = @()
        }
        $rows.Add([pscustomobject]@{
            Section               = 'Info'
            vCenter               = $vc
            Datacenter            = ''
            Cluster               = ''
            HostName              = ''
            HostOS                = ''
            HBA_Adapter           = ''
            HBA_MakeModel         = ''
            HBA_FirmwareVersion   = ''
            HBA_Driver            = ''
            HBA_QueueDepth        = ''
            HBA_LinkSpeed         = ''
            WWPN                  = ''
            IQN                   = ''
            Device_Canonical      = ''
            Device_DisplayName    = ''
            SATP                  = ''
            PSP                   = ''
            PSP_Options           = ''
            Device_PathCount      = ''
            Device_ActivePaths    = ''
            Device_Operational    = ''
            Path_RuntimeName      = ''
            Path_Target           = ''
            Path_LUN              = ''
            Path_State            = ''
            Path_Preferred        = ''
            Path_Working          = ''
            Path_Transport        = ''
            Summary_Metric        = 'Info'
            Summary_Value         = "Connection failed for $vc"
            Summary_Notes         = ''
        })
        continue
    }

    try {
        # Build filter locations (clusters/datacenters)
        $locations = @()
        if ($ClusterFilterNames.Count -gt 0) {
            foreach ($cName in $ClusterFilterNames) {
                $clusterObj = Get-Cluster -Name $cName -ErrorAction SilentlyContinue
                if ($clusterObj) { $locations += $clusterObj } else { Write-Log "Cluster not found: $cName" }
            }
        }
        if ($DatacenterFilterNames.Count -gt 0) {
            foreach ($dName in $DatacenterFilterNames) {
                $dcObj = Get-Datacenter -Name $dName -ErrorAction SilentlyContinue
                if ($dcObj) { $locations += $dcObj } else { Write-Log "Datacenter not found: $dName" }
            }
        }

        # Get hosts
        $esxiHosts = @()
        if ($locations.Count -gt 0) {
            $esxiHosts = Get-VMHost -Location $locations -ErrorAction SilentlyContinue
        } else {
            $esxiHosts = Get-VMHost -ErrorAction SilentlyContinue
        }

        # Regex filter
        $esxiHosts = $esxiHosts | Where-Object { $_.Name -match $HostNameRegex }

        if (-not $esxiHosts -or $esxiHosts.Count -eq 0) {
            Write-Warning "No hosts matched filters on $vc"
            Write-Log     "No hosts matched filters on $vc"
            $summaryByVC[$vc] = @{
                HostsProcessed     = 0
                HostsWithDeadPaths = 0
                TotalPaths         = 0
                DeadPaths          = 0
                PSPCounts          = @{}
                SATPCounts         = @{}
                HostsDeadPathsList = @()
            }
            $rows.Add([pscustomobject]@{
                Section               = 'Info'
                vCenter               = $vc
                Datacenter            = ''
                Cluster               = ''
                HostName              = ''
                HostOS                = ''
                HBA_Adapter           = ''
                HBA_MakeModel         = ''
                HBA_FirmwareVersion   = ''
                HBA_Driver            = ''
                HBA_QueueDepth        = ''
                HBA_LinkSpeed         = ''
                WWPN                  = ''
                IQN                   = ''
                Device_Canonical      = ''
                Device_DisplayName    = ''
                SATP                  = ''
                PSP                   = ''
                PSP_Options           = ''
                Device_PathCount      = ''
                Device_ActivePaths    = ''
                Device_Operational    = ''
                Path_RuntimeName      = ''
                Path_Target           = ''
                Path_LUN              = ''
                Path_State            = ''
                Path_Preferred        = ''
                Path_Working          = ''
                Path_Transport        = ''
                Summary_Metric        = 'Info'
                Summary_Value         = "No hosts matched filters for $vc"
                Summary_Notes         = ''
            })
            Disconnect-VIServer -Server $vc -Confirm:$false | Out-Null
            continue
        }

        # Init summary
        $summaryByVC[$vc] = @{
            HostsProcessed     = 0
            HostsWithDeadPaths = 0
            TotalPaths         = 0
            DeadPaths          = 0
            PSPCounts          = @{}
            SATPCounts         = @{}
            HostsDeadPathsList = @()
        }

        foreach ($esxiHost in $esxiHosts) {
            $summaryByVC[$vc].HostsProcessed++
            Write-Host "Processing host: $($esxiHost.Name)" -ForegroundColor Green
            Write-Log   "Processing host: $($esxiHost.Name)"

            $hostOs = "ESXi $($esxiHost.Version) (Build $($esxiHost.Build))"

            # Resolve cluster/datacenter names
            $clusterName = ''
            try {
                $clusterObj = Get-Cluster -VMHost $esxiHost -ErrorAction SilentlyContinue
                if ($clusterObj) { $clusterName = $clusterObj.Name }
            } catch { $clusterName = '' }

            $datacenterName = ''
            try {
                if ($clusterObj) {
                    $clusterView  = Get-View -Id $clusterObj.Id -ErrorAction SilentlyContinue
                    $dcRef        = $clusterView.Parent
                    $dcView       = Get-View -Id $dcRef -ErrorAction SilentlyContinue
                    if ($dcView -and $dcView.Name) { $datacenterName = $dcView.Name }
                }
            } catch { $datacenterName = '' }

            # ESXCLI v2
            $esxcli = Get-EsxCli -VMHost $esxiHost -V2

            # ---------- HBA info ----------
            # NOTE: Avoid FCoE in -Type to prevent enum errors; filter types afterward.
            $hbas = @()
            try {
                $hbas = Get-VMHostHba -VMHost $esxiHost -ErrorAction SilentlyContinue |
                    Where-Object { $_.Type -match 'FibreChannel|iScsi|ParallelScsi' }
            } catch {
                Write-Log "Get-VMHostHba failed on $($esxiHost.Name): $($_.Exception.Message)"
                $hbas = @()
            }

            $adapterList = Try-Esxcli { $esxcli.storage.core.adapter.list.Invoke() }
            $hbaInfoByAdapter = @{} # adapter -> @{MakeModel; Firmware; Driver; QueueDepth; LinkSpeed}

            foreach ($adapter in $adapterList) {
                $adapterName = Get-PropValue -Obj $adapter -Candidates @('Adapter','Name')
                if (-not $adapterName) { continue }

                $listDesc = Get-PropValue -Obj $adapter -Candidates @('Description')

                # adapter.get details (safe)
                $adapterDetails = Try-Esxcli {
                    $getArgs = $esxcli.storage.core.adapter.get.CreateArgs()
                    $getArgs.Adapter = $adapterName
                    $esxcli.storage.core.adapter.get.Invoke($getArgs)
                }

                $fw        = if ($adapterDetails) { Get-PropValue -Obj $adapterDetails -Candidates @('Firmware Version','FirmwareVersion') } else { '' }
                $drv       = if ($adapterDetails) { Get-PropValue -Obj $adapterDetails -Candidates @('Driver') } else { '' }
                $qDepth    = if ($adapterDetails) { Get-PropValue -Obj $adapterDetails -Candidates @('Queue Depth','QueueDepth') } else { '' }
                $linkSpeed = if ($adapterDetails) { Get-PropValue -Obj $adapterDetails -Candidates @('Link Speed','Speed') } else { '' }

                # Derive Make/Model from HBA object
                $makeModel = ''
                $hbaMatch = $hbas | Where-Object { $_.Device -eq $adapterName }
                if ($hbaMatch) {
                    $parts = @()
                    if ($hbaMatch.Manufacturer) { $parts += $hbaMatch.Manufacturer }
                    if ($hbaMatch.Model)        { $parts += $hbaMatch.Model }
                    $makeModel = ($parts -join ' ').Trim()
                } elseif ($listDesc) {
                    $makeModel = $listDesc
                }

                $hbaInfoByAdapter[$adapterName] = @{
                    MakeModel = $makeModel
                    Firmware  = $fw
                    Driver    = $drv
                    QueueDepth= $qDepth
                    LinkSpeed = $linkSpeed
                }
            }

            # ---------- FC WWPNs & iSCSI sessions ----------
            $fcList    = Try-Esxcli { $esxcli.storage.san.fc.list.Invoke() }
            $iscsiList = Try-Esxcli { $esxcli.iscsi.session.list.Invoke() }

            $wwpnByAdapter = @{}
            foreach ($fc in $fcList) {
                $fcAdapter = Get-PropValue -Obj $fc -Candidates @('Adapter')
                $portName  = Get-PropValue -Obj $fc -Candidates @('Port Name','PortName','WWPN')
                if ($fcAdapter) { $wwpnByAdapter[$fcAdapter] = $portName }
            }

            $iqnByAdapter = @{}
            foreach ($is in $iscsiList) {
                $isAdapter = Get-PropValue -Obj $is -Candidates @('Adapter')
                $targetIQN = Get-PropValue -Obj $is -Candidates @('Target Name','TargetName','IQN')
                if ($isAdapter) {
                    if (-not $iqnByAdapter.ContainsKey($isAdapter)) { $iqnByAdapter[$isAdapter] = @() }
                    if ($targetIQN) { $iqnByAdapter[$isAdapter] += $targetIQN }
                }
            }

            # ---------- Multipath device + paths ----------
            $nmpDevices = Try-Esxcli { $esxcli.storage.nmp.device.list.Invoke() }
            $nmpPaths   = Try-Esxcli { $esxcli.storage.nmp.path.list.Invoke() }
            $corePaths  = Try-Esxcli { $esxcli.storage.core.path.list.Invoke() }

            # Device map for SATP/PSP/options/Display/Operational
            $nmpDeviceMap = @{}
            foreach ($dev in $nmpDevices) {
                $canonical = Get-PropValue -Obj $dev -Candidates @('Device','Canonical Name')
                if (-not $canonical) { continue }
                $satp       = Get-PropValue -Obj $dev -Candidates @('Storage Array Type Plugin','SATP')
                $psp        = Get-PropValue -Obj $dev -Candidates @('Path Selection Policy','PSP','PathSelectionPolicy')
                $pspOptions = Get-PropValue -Obj $dev -Candidates @('Path Selection Policy Device Config','PSP Options','PathSelectionPolicyDeviceConfig')
                $dispName   = Get-PropValue -Obj $dev -Candidates @('Display Name','Device Display Name','Name')
                $opState    = Get-PropValue -Obj $dev -Candidates @('Operational State','OperationalState')

                $nmpDeviceMap[$canonical] = @{
                    SATP       = $satp
                    PSP        = $psp
                    PSPOptions = $pspOptions
                    Display    = $dispName
                    OpState    = $opState
                }

                # Summary counts
                if ($satp) {
                    if (-not $summaryByVC[$vc].SATPCounts.ContainsKey($satp)) { $summaryByVC[$vc].SATPCounts[$satp] = 0 }
                    $summaryByVC[$vc].SATPCounts[$satp]++
                }
                if ($psp) {
                    if (-not $summaryByVC[$vc].PSPCounts.ContainsKey($psp)) { $summaryByVC[$vc].PSPCounts[$psp] = 0 }
                    $summaryByVC[$vc].PSPCounts[$psp]++
                }
            }

            # Path counts per device
            $pathCountByDevice   = @{}
            $activeCountByDevice = @{}
            $hostHasDeadPath     = $false

            # NMP runtime meta
            $nmpPathMetaByRuntime = @{}
            foreach ($np in $nmpPaths) {
                $rt   = Get-PropValue -Obj $np -Candidates @('Runtime Name','RuntimeName','Name','Path')
                if (-not $rt) { continue }
                $work = Get-PropValue -Obj $np -Candidates @('Working','IsWorking','Is Working Path')
                $pref = Get-PropValue -Obj $np -Candidates @('Preferred')
                $stat = Get-PropValue -Obj $np -Candidates @('State','Path State')
                $nmpPathMetaByRuntime[$rt] = @{ Working=$work; Preferred=$pref; State=$stat }
            }

            foreach ($p in $corePaths) {
                $dev      = Get-PropValue -Obj $p -Candidates @('Device')
                if (-not $dev) { continue }
                if (-not $pathCountByDevice.ContainsKey($dev)) { $pathCountByDevice[$dev] = 0 }
                $pathCountByDevice[$dev]++

                $state    = Get-PropValue -Obj $p -Candidates @('State','Path State')
                if ($state -and ($state.ToString().ToLower().Contains('dead'))) {
                    $hostHasDeadPath = $true
                    $summaryByVC[$vc].DeadPaths++
                }
                $summaryByVC[$vc].TotalPaths++

                $rt = Get-PropValue -Obj $p -Candidates @('Runtime Name','RuntimeName','Name','Path')
                $workFlag = $null
                if ($rt -and $nmpPathMetaByRuntime.ContainsKey($rt)) {
                    $workFlag = $nmpPathMetaByRuntime[$rt].Working
                }
                $isActive = $false
                if ($state -and ($state.ToString().ToLower().Contains('active'))) { $isActive = $true }
                if ($workFlag -and $workFlag.ToString().ToLower() -in @('true','y','yes','working')) { $isActive = $true }

                if ($isActive) {
                    if (-not $activeCountByDevice.ContainsKey($dev)) { $activeCountByDevice[$dev] = 0 }
                    $activeCountByDevice[$dev]++
                }
            }

            if ($hostHasDeadPath) {
                $summaryByVC[$vc].HostsWithDeadPaths++
                $summaryByVC[$vc].HostsDeadPathsList += $esxiHost.Name
            }

            # ---------- Emit detailed rows (one per path); else device-only rows ----------
            if ($corePaths -and $corePaths.Count -gt 0) {
                foreach ($p in $corePaths) {
                    $canonical = Get-PropValue -Obj $p -Candidates @('Device')
                    $rtName    = Get-PropValue -Obj $p -Candidates @('Runtime Name','RuntimeName','Name','Path')
                    $adapter   = Get-PropValue -Obj $p -Candidates @('Adapter')
                    $target    = Get-PropValue -Obj $p -Candidates @('Target','Target Number')
                    $lun       = Get-PropValue -Obj $p -Candidates @('LUN','Lun','Lun Id')
                    $state     = Get-PropValue -Obj $p -Candidates @('State','Path State')
                    $pref      = Get-PropValue -Obj $p -Candidates @('Preferred')
                    $transport = Get-PropValue -Obj $p -Candidates @('Transport','Transport Type')

                    if ($rtName -and $nmpPathMetaByRuntime.ContainsKey($rtName)) {
                        if (-not $pref)  { $pref  = $nmpPathMetaByRuntime[$rtName].Preferred }
                        if (-not $state) { $state = $nmpPathMetaByRuntime[$rtName].State }
                    }

                    $hbaInfo   = if ($adapter -and $hbaInfoByAdapter.ContainsKey($adapter)) { $hbaInfoByAdapter[$adapter] } else { @{} }
                    $makeModel = if ($hbaInfo.ContainsKey('MakeModel')) { $hbaInfo.MakeModel } else { '' }
                    $fw        = if ($hbaInfo.ContainsKey('Firmware'))  { $hbaInfo.Firmware }  else { '' }
                    $drv       = if ($hbaInfo.ContainsKey('Driver'))    { $hbaInfo.Driver }    else { '' }
                    $qDepth    = if ($hbaInfo.ContainsKey('QueueDepth')){ $hbaInfo.QueueDepth }else { '' }
                    $linkSpeed = if ($hbaInfo.ContainsKey('LinkSpeed')) { $hbaInfo.LinkSpeed } else { '' }

                    $satp      = ''
                    $psp       = ''
                    $pspOpt    = ''
                    $dispName  = ''
                    $opState   = ''
                    if ($canonical -and $nmpDeviceMap.ContainsKey($canonical)) {
                        $satp     = $nmpDeviceMap[$canonical].SATP
                        $psp      = $nmpDeviceMap[$canonical].PSP
                        $pspOpt   = $nmpDeviceMap[$canonical].PSPOptions
                        $dispName = $nmpDeviceMap[$canonical].Display
                        $opState  = $nmpDeviceMap[$canonical].OpState
                    }

                    $totalPaths  = if ($canonical -and $pathCountByDevice.ContainsKey($canonical))   { $pathCountByDevice[$canonical] } else { $null }
                    $activePaths = if ($canonical -and $activeCountByDevice.ContainsKey($canonical)) { $activeCountByDevice[$canonical] } else { $null }

                    # WWPN/IQN
                    $wwpn = if ($adapter -and $wwpnByAdapter.ContainsKey($adapter)) { $wwpnByAdapter[$adapter] } else { '' }
                    $iqn  = if ($adapter -and $iqnByAdapter.ContainsKey($adapter))  { ($iqnByAdapter[$adapter] -join '; ') } else { '' }

                    $rows.Add([pscustomobject]@{
                        Section               = 'Path'
                        vCenter               = $vc
                        Datacenter            = $datacenterName
                        Cluster               = $clusterName
                        HostName              = $esxiHost.Name
                        HostOS                = $hostOs
                        HBA_Adapter           = $adapter
                        HBA_MakeModel         = $makeModel
                        HBA_FirmwareVersion   = $fw
                        HBA_Driver            = $drv
                        HBA_QueueDepth        = $qDepth
                        HBA_LinkSpeed         = $linkSpeed
                        WWPN                  = $wwpn
                        IQN                   = $iqn
                        Device_Canonical      = $canonical
                        Device_DisplayName    = $dispName
                        SATP                  = $satp
                        PSP                   = $psp
                        PSP_Options           = $pspOpt
                        Device_PathCount      = $totalPaths
                        Device_ActivePaths    = $activePaths
                        Device_Operational    = $opState
                        Path_RuntimeName      = $rtName
                        Path_Target           = $target
                        Path_LUN              = $lun
                        Path_State            = $state
                        Path_Preferred        = $pref
                        Path_Working          = if ($rtName -and $nmpPathMetaByRuntime.ContainsKey($rtName)) { $nmpPathMetaByRuntime[$rtName].Working } else { '' }
                        Path_Transport        = $transport
                        Summary_Metric        = ''
                        Summary_Value         = ''
                        Summary_Notes         = ''
                    })
                }
            } else {
                # No core paths; emit device-only rows
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

                    $rows.Add([pscustomobject]@{
                        Section               = 'Device'
                        vCenter               = $vc
                        Datacenter            = $datacenterName
                        Cluster               = $clusterName
                        HostName              = $esxiHost.Name
                        HostOS                = $hostOs
                        HBA_Adapter           = ''
                        HBA_MakeModel         = ''
                        HBA_FirmwareVersion   = ''
                        HBA_Driver            = ''
                        HBA_QueueDepth        = ''
                        HBA_LinkSpeed         = ''
                        WWPN                  = ''
                        IQN                   = ''
                        Device_Canonical      = $canonical
                        Device_DisplayName    = $dispName
                        SATP                  = $satp
                        PSP                   = $psp
                        PSP_Options           = $pspOpt
                        Device_PathCount      = $totalPaths
                        Device_ActivePaths    = $activePaths
                        Device_Operational    = $opState
                        Path_RuntimeName      = ''
                        Path_Target           = ''
                        Path_LUN              = ''
                        Path_State            = ''
                        Path_Preferred        = ''
                        Path_Working          = ''
                        Path_Transport        = ''
                        Summary_Metric        = ''
                        Summary_Value         = ''
                        Summary_Notes         = ''
                    })
                }
            }

        } # end foreach host

        # Append summary rows for this vCenter
        $rows.Add([pscustomobject]@{
            Section             = 'Summary'
            vCenter             = $vc
            Datacenter          = ''
            Cluster             = ''
            HostName            = ''
            HostOS              = ''
            HBA_Adapter         = ''
            HBA_MakeModel       = ''
            HBA_FirmwareVersion = ''
            HBA_Driver          = ''
            HBA_QueueDepth      = ''
            HBA_LinkSpeed       = ''
            WWPN                = ''
            IQN                 = ''
            Device_Canonical    = ''
            Device_DisplayName  = ''
            SATP                = ''
            PSP                 = ''
            PSP_Options         = ''
            Device_PathCount    = ''
            Device_ActivePaths  = ''
            Device_Operational  = ''
            Path_RuntimeName    = ''
            Path_Target         = ''
            Path_LUN            = ''
            Path_State          = ''
            Path_Preferred      = ''
            Path_Working        = ''
            Path_Transport      = ''
            Summary_Metric      = 'vCenter Summary'
            Summary_Value       = "HostsProcessed=$($summaryByVC[$vc].HostsProcessed); HostsWithDeadPaths=$($summaryByVC[$vc].HostsWithDeadPaths); TotalPaths=$($summaryByVC[$vc].TotalPaths); DeadPaths=$($summaryByVC[$vc].DeadPaths)"
            Summary_Notes       = "HostsWithDeadPathsList=$([string]::Join('; ', $summaryByVC[$vc].HostsDeadPathsList))"
        })

        foreach ($pspKey in $summaryByVC[$vc].PSPCounts.Keys) {
            $rows.Add([pscustomobject]@{
                Section             = 'Summary'
                vCenter             = $vc
                Datacenter          = ''
                Cluster             = ''
                HostName            = ''
                HostOS              = ''
                HBA_Adapter         = ''
                HBA_MakeModel       = ''
                HBA_FirmwareVersion = ''
                HBA_Driver          = ''
                HBA_QueueDepth      = ''
                HBA_LinkSpeed       = ''
                WWPN                = ''
                IQN                 = ''
                Device_Canonical    = ''
                Device_DisplayName  = ''
                SATP                = ''
                PSP                 = $pspKey
                PSP_Options         = ''
                Device_PathCount    = ''
                Device_ActivePaths  = ''
                Device_Operational  = ''
                Path_RuntimeName    = ''
                Path_Target         = ''
                Path_LUN            = ''
                Path_State          = ''
                Path_Preferred      = ''
                Path_Working        = ''
                Path_Transport      = ''
                Summary_Metric      = 'PSP Count'
                Summary_Value       = $summaryByVC[$vc].PSPCounts[$pspKey]
                Summary_Notes       = ''
            })
        }

        foreach ($satpKey in $summaryByVC[$vc].SATPCounts.Keys) {
            $rows.Add([pscustomobject]@{
                Section             = 'Summary'
                vCenter             = $vc
                Datacenter          = ''
                Cluster             = ''
                HostName            = ''
                HostOS              = ''
                HBA_Adapter         = ''
                HBA_MakeModel       = ''
                HBA_FirmwareVersion = ''
                HBA_Driver          = ''
                HBA_QueueDepth      = ''
                HBA_LinkSpeed       = ''
                WWPN                = ''
                IQN                 = ''
                Device_Canonical    = ''
                Device_DisplayName  = ''
                SATP                = $satpKey
                PSP                 = ''
                PSP_Options         = ''
                Device_PathCount    = ''
                Device_ActivePaths  = ''
                Device_Operational  = ''
                Path_RuntimeName    = ''
                Path_Target         = ''
                Path_LUN            = ''
                Path_State          = ''
                Path_Preferred      = ''
                Path_Working        = ''
                Path_Transport      = ''
                Summary_Metric      = 'SATP Count'
                Summary_Value       = $summaryByVC[$vc].SATPCounts[$satpKey]
                Summary_Notes       = ''
            })
        }

    } finally {
        Disconnect-VIServer -Server $vc -Confirm:$false | Out-Null
        Write-Log "Disconnected from $vc"
    }
} # end foreach vCenter

# Ensure at least one row so CSV is not empty
if ($rows.Count -eq 0) {
    $rows.Add([pscustomobject]@{
        Section               = 'Info'
        vCenter               = ''
        Datacenter            = ''
        Cluster               = ''
        HostName              = ''
        HostOS                = ''
        HBA_Adapter           = ''
        HBA_MakeModel         = ''
        HBA_FirmwareVersion   = ''
        HBA_Driver            = ''
        HBA_QueueDepth        = ''
        HBA_LinkSpeed         = ''
        WWPN                  = ''
        IQN                   = ''
        Device_Canonical      = ''
        Device_DisplayName    = ''
        SATP                  = ''
        PSP                   = ''
        PSP_Options           = ''
        Device_PathCount      = ''
        Device_ActivePaths    = ''
        Device_Operational    = ''
        Path_RuntimeName      = ''
        Path_Target           = ''
        Path_LUN              = ''
        Path_State            = ''
        Path_Preferred        = ''
        Path_Working          = ''
        Path_Transport        = ''
        Summary_Metric        = 'Info'
        Summary_Value         = 'No data collected; verify credentials, permissions, filters, and connectivity.'
        Summary_Notes         = ''
    })
    Write-Log "No data rows collected; wrote informational row"
}

# Export CSV
try {
    $rows | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8 -Force
    Write-Host "All results saved to: $OutputCsv" -ForegroundColor Yellow
    Write-Log   "Exported CSV to $OutputCsv with $($rows.Count) rows"
} catch {
    Write-Error "Failed to export CSV: $($_.Exception.Message)"
    Write-Log   "Export CSV failed: $($_.Exception.Message)"
}
