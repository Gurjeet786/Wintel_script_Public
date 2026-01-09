
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
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }

Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false | Out-Null

# Regex filter for hosts (e.g., only prod)
$HostFilterRegex = 'prod'   # Change as needed
$ClusterFilter   = ''       # Optional cluster name filter

# ---------------------------
# Helper for property lookup
# ---------------------------
function Get-PropValue {
    param($Obj,[string[]]$Candidates)
    if (-not $Obj) { return $null }
    $map=@{}
    foreach ($p in $Obj.PSObject.Properties){$map[($p.Name -replace '[\s_-]','').ToLower()]=$p.Name}
    foreach ($cand in $Candidates){
        $key=($cand -replace '[\s_-]','').ToLower()
        if ($map.ContainsKey($key)){return $Obj.($map[$key])}
    }
    return $null
}

# ---------------------------
# Data collection
# ---------------------------
$rows = New-Object System.Collections.Generic.List[object]
$summary = @{}

foreach ($vc in $VCServers) {
    Write-Host "Connecting to vCenter $vc ..." -ForegroundColor Cyan
    $vcCred = Get-Credential -Message "Enter credentials for $vc"
    try {
        Connect-VIServer -Server $vc -Credential $vcCred -ErrorAction Stop | Out-Null
    } catch {
        Write-Warning "Failed to connect to $vc: $($_.Exception.Message)"
        continue
    }

    $esxiHosts = Get-VMHost | Where-Object {
        ($_ | Select-Object -ExpandProperty Name) -match $HostFilterRegex -and
        ($ClusterFilter -eq '' -or $_.Parent.Name -eq $ClusterFilter)
    }

    foreach ($esxiHost in $esxiHosts) {
        Write-Host "Processing host: $($esxiHost.Name)" -ForegroundColor Green
        $hostOs = "ESXi $($esxiHost.Version) (Build $($esxiHost.Build))"
        $esxcli = Get-EsxCli -VMHost $esxiHost -V2

        # HBA details
        $hbas = Get-VMHostHba -VMHost $esxiHost -Type FibreChannel,iSCSI -ErrorAction SilentlyContinue
        $adapterList = $esxcli.storage.core.adapter.list.Invoke()
        $firmwareByAdapter=@{}
        foreach ($adapter in $adapterList){
            $adapterName=Get-PropValue $adapter @('Adapter','Name')
            $getArgs=$esxcli.storage.core.adapter.get.CreateArgs()
            $getArgs.Adapter=$adapterName
            $adapterDetails=$esxcli.storage.core.adapter.get.Invoke($getArgs)
            $fw=Get-PropValue $adapterDetails @('Firmware Version','FirmwareVersion')
            $drv=Get-PropValue $adapterDetails @('Driver')
            $qDepth=Get-PropValue $adapterDetails @('Queue Depth')
            $linkSpeed=Get-PropValue $adapterDetails @('Link Speed')
            $firmwareByAdapter[$adapterName]=@{FW=$fw;Driver=$drv;QueueDepth=$qDepth;LinkSpeed=$linkSpeed}
        }

        # FC WWPNs
        $fcList=@()
        try {$fcList=$esxcli.storage.san.fc.list.Invoke()} catch {}

        # iSCSI sessions
        $iscsiList=@()
        try {$iscsiList=$esxcli.iscsi.session.list.Invoke()} catch {}

        # Multipath info
        $nmpDevices=$esxcli.storage.nmp.device.list.Invoke()
        $corePaths=$esxcli.storage.core.path.list.Invoke()

        foreach ($p in $corePaths){
            $canonical=Get-PropValue $p @('Device')
            $rtName=Get-PropValue $p @('Runtime Name')
            $adapter=Get-PropValue $p @('Adapter')
            $state=Get-PropValue $p @('State')
            $target=Get-PropValue $p @('Target')
            $lun=Get-PropValue $p @('LUN')
            $fw=$firmwareByAdapter[$adapter].FW
            $drv=$firmwareByAdapter[$adapter].Driver
            $qDepth=$firmwareByAdapter[$adapter].QueueDepth
            $linkSpeed=$firmwareByAdapter[$adapter].LinkSpeed

            # FC WWPN mapping
            $wwpn=''
            foreach ($fc in $fcList){
                if ((Get-PropValue $fc @('Adapter')) -eq $adapter){
                    $wwpn=Get-PropValue $fc @('Port Name')
                }
            }

            # iSCSI mapping
            $iqn=''
            foreach ($iscsi in $iscsiList){
                if ((Get-PropValue $iscsi @('Adapter')) -eq $adapter){
                    $iqn=Get-PropValue $iscsi @('Target Name')
                }
            }

            $rows.Add([pscustomobject]@{
                vCenter=$vc;HostName=$esxiHost.Name;HostOS=$hostOs
                HBA_Adapter=$adapter;Firmware=$fw;Driver=$drv;QueueDepth=$qDepth;LinkSpeed=$linkSpeed
                WWPN=$wwpn;IQN=$iqn
                Device=$canonical;Path_RuntimeName=$rtName;Path_State=$state;Target=$target;LUN=$lun
            })

            # Summary counts
            if (-not $summary.ContainsKey($vc)){$summary[$vc]=@{DeadPaths=0;TotalPaths=0}}
            $summary[$vc].TotalPaths++
            if ($state -match 'dead'){$summary[$vc].DeadPaths++}
        }
    }

    Disconnect-VIServer -Server $vc -Confirm:$false | Out-Null
}

# Export all rows
$rows | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
Write-Host "All results saved to: $OutputCsv" -ForegroundColor Yellow

# Summary output
Write-Host "`nSummary Report:" -ForegroundColor Cyan
foreach ($vc in $summary.Keys){
    Write-Host "$vc : Total Paths=$($summary[$vc].TotalPaths), Dead Paths=$($summary[$vc].DeadPaths)"
}
