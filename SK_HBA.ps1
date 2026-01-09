
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
$LogFile      = Join-Path $OutputFolder 'VMware_HBA_Multipath.log'

# Filters (leave empty/blank to disable)
$ClusterFilterNames   = @()        # e.g., @('Prod-Cluster-A','Prod-Cluster-B')
$DatacenterFilterNames= @()        # e.g., @('DC-East','DC-West')
$HostNameRegex        = '.*'       # e.g., '^.*prod.*$' ; default matches all

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

function Safe-InvokeEsxcli {
    param(
        [Parameter(Mandatory=$true)]$Method
    )
    try {
        return $Method.Invoke()
    } catch {
        Write-Log "ESXCLI invoke failed: $($_.Exception.Message)"
        return @()
    }
}

# ---------------------------
# Data collection structures
# ---------------------------
$rows = New-Object System.Collections.Generic.List[object]

# Summary per-vCenter
$summaryByVC = @{} # vc -> @{HostsProcessed=0; HostsWithDeadPaths=0; TotalPaths=0; DeadPaths=0; PSPCounts=@{}; SATPCounts=@{} }

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
        # Record summary stub and continue
        $summaryByVC[$vc] = @{
            HostsProcessed    = 0
            HostsWithDeadPaths= 0
            TotalPaths        = 0
            DeadPaths         = 0
            PSPCounts         = @{}
            SATPCounts        = @{}
        }
        continue
    }

    try {
        # Build location filter (clusters/datacenters)
        $locations = @()
        if ($ClusterFilterNames -and $ClusterFilterNames.Count -gt 0) {
            foreach ($c in $ClusterFilterNames) {
                $clusterObj = Get-Cluster -Name $c -ErrorAction SilentlyContinue
                if ($clusterObj) { $locations += $clusterObj } else { Write-Log "Cluster not found: $c" }
            }
        }
        if ($DatacenterFilterNames -and $DatacenterFilterNames.Count -gt 0) {
            foreach ($d in $DatacenterFilterNames) {
                $dcObj = Get-Datacenter -Name $d -ErrorAction SilentlyContinue
                if ($dcObj) { $locations += $dcObj } else { Write-Log "Datacenter not found: $d" }
            }
        }

        $esxiHosts = @()
        if ($locations -and $locations.Count -gt 0) {
            $esxiHosts = Get-VMHost -Location $locations -ErrorAction SilentlyContinue
        } else {
            $esxiHosts = Get-VMHost -ErrorAction SilentlyContinue
        }

        # Final regex filter on host names
        $esxiHosts = $esxiHosts | Where-Object { $_.Name -match $HostNameRegex }

        if (-not $esxiHosts -or $esxiHosts.Count -eq 0) {
            Write-Warning "No hosts matched filters on $vc"
            Write-Log     "No hosts matched filters on $vc"
            # Still append an informational row so the CSV is not empty
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
            # Initialize summary
            $summaryByVC[$vc] = @{
                HostsProcessed    = 0
                HostsWithDeadPaths= 0
                TotalPaths        = 0
                DeadPaths         = 0
                PSPCounts         = @{}
                SATPCounts        = @{}
            }
            # Move to next vCenter
            Disconnect-VIServer -Server $vc -Confirm:$false | Out-Null
            continue
        }

        # Initialize summary counters
        $summaryByVC[$vc] = @{
            HostsProcessed    = 0
            HostsWithDeadPaths= 0
            TotalPaths        = 0
            DeadPaths         = 0
            PSPCounts         = @{}
            SATPCounts        = @{}
        }

        foreach ($esxiHost in $esxiHosts) {
            $summaryByVC[$vc].HostsProcessed++
            Write-Host "Processing host: $($esxiHost.Name)" -ForegroundColor Green
            Write-Log   "Processing host: $($esxiHost.Name)"

            $hostOs = "ESXi $($esxiHost.Version) (Build $($esxiHost.Build))"

            # Resolve cluster/datacenter names (safe approach)
            $clusterName  = ($esxiHost | Get-Cluster -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name)
            if (-not $clusterName) { $clusterName = '' }
            $datacenterName = ''
            try {
                $parentCluster = ($esxiHost | Get-Cluster -ErrorAction SilentlyContinue)
                if ($parentCluster) {
                    $dc = ($parentCluster | Get-View).Parent
                    if ($dc) {
                        $dcView = Get-View -Id $dc -ErrorAction SilentlyContinue
                        if ($dcView.Name) { $datacenterName = $dcView.Name }
                    }
                }
            } catch { $datacenterName = '' }

            $esxcli = Get-EsxCli -VMHost $esxiHost -V2

            # ---------- HBA info & firmware/driver/queue/link ----------
            # NOTE: FCoE type removed to avoid enum errors in some PowerCLI versions.
            $hbas = @()
            try {
                $hbas = Get-VMHostHba -VMHost $esxiHost -Type FibreChannel,iSCSI -ErrorAction SilentlyContinue
            } catch {
                Write-Log "Get-VMHostHba failed on $($esxiHost.Name): $($_.Exception.Message)"
                $hbas = @()
            }

            $adapterList = Safe-InvokeEsxcli -Method $esxcli.storage.core.adapter.list
            $hbaInfoByAdapter = @{} # adapter -> @{MakeModel; Firmware; Driver; QueueDepth; LinkSpeed}

            foreach ($adapter in $adapterList) {
                $adapterName = Get-PropValue -Obj $adapter -Candidates @('Adapter','Name')
                if (-not $adapterName) { continue }

                # Pull "Description" from list; fallback later if HBA object exists
                $listDesc = Get-PropValue -Obj $adapter -Candidates @('Description')

                # Details via get
                $getArgs = $esxcli.storage.core.adapter.get.CreateArgs()
                $getArgs.Adapter = $adapterName
                $adapterDetails  = Safe-InvokeEsxcli -Method ($esxcli.storage.core.adapter.get.CreateArgs().GetType() | Out-Null; $esxcli.storage.core.adapter.get.Invoke($getArgs))
                # Above line ensures the CreateArgs exists; if Invoke fails, adapterDetails becomes @()

                # Because Invoke in try/catch is tricky, retrieve via a safer pattern:
                try {
                    $getArgs = $esxcli.storage.core.adapter.get.CreateArgs()
                    $getArgs.Adapter = $adapterName
                    $adapterDetails = $esxcli.storage.core.adapter.get.Invoke($getArgs)
                } catch {
                    $adapterDetails = $null
                    Write-Log "adapter.get failed for $adapterName: $($_.Exception.Message)"
                }

                $fw        = if ($adapterDetails) { Get-PropValue -Obj $adapterDetails -Candidates @('Firmware Version','FirmwareVersion') } else { '' }
                $drv       = if ($adapterDetails) { Get-PropValue -Obj $adapterDetails -Candidates @('Driver') } else { '' }
                $qDepth    = if ($adapterDetails) { Get-PropValue -Obj $adapterDetails -Candidates @('Queue Depth','QueueDepth') } else { '' }
                $linkSpeed = if ($adapterDetails) { Get-PropValue -Obj $adapterDetails -Candidates @('Link Speed','Speed') } else { '' }

                # Attempt to derive Make/Model from PowerCLI HBA objects
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
            $fcList    = Safe-InvokeEsxcli -Method $esxcli.storage.san.fc.list
            $iscsiList = Safe-InvokeEsxcli -Method $esxcli.iscsi.session.list

            # Build maps for quick lookup
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
            $nmpDevices = Safe-InvokeEsxcli -Method $esxcli.storage.nmp.device.list
            $nmpPaths   = Safe-InvokeEsxcli -Method $esxcli.storage.nmp.path.list
            $corePaths  = Safe-InvokeEsxcli -Method $esxcli.storage.core.path.list

            # Build device map for SATP/PSP/options/Display/Operational
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

            # Build NMP runtime meta
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

            if ($hostHasDeadPath) { $summaryByVC[$vc].HostsWithDeadPaths++ }

            # ---------- Emit detailed rows (one per path); if no paths, emit device-only rows ----------
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
                # No core paths; emit device-only lines
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
            Summary_Notes       = ''
        })

        foreach ($k in $summaryByVC[$vc].PSPCounts.Keys) {
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
                PSP                 = $k
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
                Summary_Value       = $summaryByVC[$vc].PSPCounts[$k]
                Summary_Notes       = ''
            })
        }

        foreach ($k in $summaryByVC[$vc].SATPCounts.Keys) {
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
                SATP                = $k
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
                Summary_Value       = $summaryByVC[$vc].SATPCounts[$k]
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
