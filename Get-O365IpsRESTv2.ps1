<#
    ############################################################################
    The sample scripts are not supported under any Microsoft standard support
    program or service. The sample scripts are provided AS IS without warranty
    of any kind. Microsoft further disclaims all implied warranties including,
    without limitation, any implied warranties of merchantability or of fitness
    for a particular purpose. The entire risk arising out of the use or
    performance of the sample scripts and documentation remains with you. In no
    event shall Microsoft, its authors, or anyone else involved in the creation,
    production, or delivery of the scripts be liable for any damages whatsoever
    (including, without limitation, damages for loss of business profits,
    business interruption, loss of business information, or other pecuniary
    loss) arising out of the use of or inability to use the sample scripts or
    documentation, even if Microsoft has been advised of the possibility of such
    damages.
    ############################################################################
#>

# webservice root URL
$ws = "https://endpoints.office.com"

# path where client ID and latest version number will be stored
$datapath =  ".\O365IP_Script_settings.txt" #this is where settings are saved.

# fetch settings if data file exists; otherwise create new file
if (Test-Path $datapath) {
    $content = Get-Content $datapath
    $clientRequestId = $content[0]
    $tenantName = $content[1]
    $outputdir = $content[2]
    $version = $content[3]
}
else {
    Write-Host "Settings file not found. Generating new client id." -ForegroundColor Yellow
    $clientRequestId = [GUID]::NewGuid().Guid
    $tenantName = Read-Host "Enter tenant name"
    $Outputdir = Read-Host "Enter output directory"
    $version="0000000000"
    @($clientRequestId, $tenantName,$outputdir,$version) | Out-File $datapath
}

$outputpath = $outputdir + "\Deltaoutput" + $(Get-Date -Format yyMMdd-hhmm) + ".xlsx" #this is where output is saved.

function Get-BroadcastAddress{
param(

    [Parameter(Mandatory = $true)]
	$ip,
    [Parameter(Mandatory = $false)]
	$cidr,
    [Parameter(Mandatory = $false)]
	$mask
)

    $ip = [Net.IPAddress]::Parse($ip)
    $bytes = $ip.GetAddressBytes()


    if($cidr){ #cidr provided
        [int]$counter = $cidr
        $maskbytes = @()
        #build mask
        foreach($octet in $bytes){
            #get 8 bits
            $maskOctet = ""
                for ($i=0;$i -lt 8; $i++){
                    if($counter -gt 0){
                        $maskOctet = $maskOctet + "1"
                        $counter--
                    }
                    else{
                        $maskOctet = $maskOctet + "0"
                    }
                }
            $maskbytes += [convert]::ToByte($maskOctet,2)
        }
    }


    elseif($mask){ #mask provided instead
        $mask = [Net.IPAddress]::Parse($mask)
        $maskbytes = $mask.GetAddressBytes()
    }


    $broadcastbytes = @()
    for($i=0;$i -lt $bytes.count; $i++){
        $broadcastbytes += 255 -bxor $maskbytes[$i] -bor $bytes[$i]
    }

    if ($broadcastbytes.count -eq 4){
        $broadcastip = $broadcastbytes -join (".")
    }

    elseif ($broadcastbytes.count -eq 16){
        $broadcastip = ""
        for($i=0;$i -lt $broadcastbytes.count; $i++){
            $temp = $([convert]::ToString($broadcastbytes[$i],16))
            if ($temp.length -eq 1){$temp = "0" + $temp}
            if ($temp.length -eq 0){$temp = "00"}
            $broadcastip =$broadcastip + $temp
            if ($i % 2 -ne 0){$broadcastip =$broadcastip + ":"}
        }
        $broadcastip = $broadcastip.TrimEnd(":")
    }
$broadcastip = [Net.IPAddress]::Parse($broadcastip)
return @($ip.IPAddressToString,$broadcastip.ipaddresstostring)
}

$EndpointSets = Invoke-RestMethod -Uri ($ws + "/endpoints/Worldwide?TenantName=$tenantName&clientRequestId=" + $clientRequestId)


#region get full ip and url list
[array]$export = ""
[array]$urlexport = ""
#list of ips
    $iplist = $EndpointSets |?{$_.ips -notlike ""}
    foreach($line in $iplist){
        foreach ($ip in $line.ips){
            $ipinfo = $ip -split "/"
            $Ipinfo = Get-BroadcastAddress -ip $IpInfo[0] -cidr $ipinfo[1]
            $port = (((("TCP-" + (($line.tcpports -split (",") | select -Unique) -join (",TCP-") ) + "," +"UDP-" + (($line.udpports -split (",") | select -Unique) -join (",UDP-") ) + ",").TrimEnd(",")).TrimEnd("TCP-")).TrimEnd("UDP-")).TrimEnd(",")
            $export += $ip | select @{n='Service';e={$line.serviceAreaDisplayName}},@{n='CIDR';e={$ip}},@{n='StartIPRange';e={$ipinfo[0]}},@{n='EndIPRange';e={$ipinfo[1]}},@{n='Category';e={$line.category}},@{n='ExpressRoute';e={$line.expressRoute}},@{n='Required';e={$line.required}},@{n='Port';e={$port}},@{n='Notes';e={$line.notes}}
        }
    }
    $export | Export-Excel $outputpath -WorkSheetname "FullIPList"

#list of urls
    $fulllist = $EndpointSets |?{$_.urls -notlike ""}
    foreach($line in $fulllist){
        foreach ($url in $line.urls){
            $urlexport += $url | select @{n='Service';e={$line.serviceAreaDisplayName}},@{n='URL';e={$url}},@{n='Category';e={$line.category}},@{n='ExpressRoute';e={$line.expressRoute}},@{n='Required';e={$line.required}},@{n='Notes';e={$line.notes}}
        }
    }
    $urlexport | Export-Excel $outputpath -WorkSheetname "FullURLList"


$Changes = Invoke-RestMethod -Uri ($ws + "/changes/Worldwide/$version" + "?TenantName=$tenantName&clientRequestId=" + $clientRequestId)

[array]$adds = $changes | ?{$_.add -notlike ""}
[array]$removes = $changes | ?{$_.remove -notlike ""}

#parse changes
[array]$export = ""
[array]$exporturls = ""
    foreach($add in $adds){
        $line = $EndpointSets |?{$_.id -eq $add.endpointsetid}
        foreach($ip in $add.add.ips){
            $IpInfo = $ip -split "/" 
            $Ipinfo = Get-BroadcastAddress -ip $IpInfo[0] -cidr $ipinfo[1]
            $port = (((("TCP-" + (($line.tcpports -split (",") | select -Unique) -join (",TCP-") ) + "," +"UDP-" + (($line.udpports -split (",") | select -Unique) -join (",UDP-") ) + ",").TrimEnd(",")).TrimEnd("TCP-")).TrimEnd("UDP-")).TrimEnd(",")
            $export += $ip | select @{n='Service';e={$line.serviceAreaDisplayName}},@{n='CIDR';e={$ip}},@{n='StartIPRange';e={$ipinfo[0]}},@{n='EndIPRange';e={$ipinfo[1]}},@{n='Category';e={$line.category}},@{n='ExpressRoute';e={$line.expressRoute}},@{n='required';e={$line.required}},@{n='Port';e={$port}},@{n='EffectiveDate';e={$add.add.effectiveDate}}
        foreach($url in $add.add.urls){
            $exporturls += $url,$add.add.effectiveDate
            }
        }
    }
    $export | Sort-Object effectivedate| Export-Excel $outputpath -WorkSheetname "AddIPList"
    $exporturls | Sort-Object effectivedate| Export-Excel $outputpath -WorkSheetname "AddURLList"

[array]$export = ""
[array]$exporturls = ""
    foreach($remove in $removes){
        $line = $EndpointSets |?{$_.id -eq $remove.endpointsetid}
        foreach($ip in $remove.remove.ips){
            if($EndpointSets.ips -notcontains $ip){
                $IpInfo = $ip -split "/" 
                $Ipinfo = Get-BroadcastAddress -ip $IpInfo[0] -cidr $ipinfo[1]
                $port = (((("TCP-" + (($line.tcpports -split (",") | select -Unique) -join (",TCP-") ) + "," +"UDP-" + (($line.udpports -split (",") | select -Unique) -join (",UDP-") ) + ",").TrimEnd(",")).TrimEnd("TCP-")).TrimEnd("UDP-")).TrimEnd(",")
                $export += $ip | select @{n='Service';e={$line.serviceAreaDisplayName}},@{n='CIDR';e={$ip}},@{n='StartIPRange';e={$ipinfo[0]}},@{n='EndIPRange';e={$ipinfo[1]}},@{n='Category';e={$line.category}},@{n='ExpressRoute';e={$line.expressRoute}},@{n='required';e={$line.required}},@{n='Port';e={$port}},@{n='EffectiveDate';e={$add.add.effectiveDate}}
            }
        }        
        foreach($url in $remove.remove.urls){
            $exporturls += $url,$remove.remove.effectivedate
        }
    }
    $export | Sort-Object effectivedate| Export-Excel $outputpath -WorkSheetname "RemoveIPList"
    $exporturls | Sort-Object effectivedate| Export-Excel $outputpath -WorkSheetname "RemoveURLList"


#get current version and save to file.
$version= (Invoke-RestMethod -Uri ($ws + "/version/Worldwide?TenantName=$tenantName&clientRequestId=" + $clientRequestId)).latest
@($clientRequestId, $tenantName,$outputdir,$version) | Out-File $datapath

#open file and exit
Read-Host "Press enter to open output file"
explorer $outputpath
