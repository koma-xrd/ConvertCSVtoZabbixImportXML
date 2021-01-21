$computers = Import-Csv H:\hardware\combined.csv
$xmlPath = "H:\hardware\finalconvert.xml"

#!!! after combining nl_Hardware_D + nl_Port-MAC-IP_D in EXCEL power-query
# export as csv in UTF8 Format!!! with following columns
#typ,hwid,serial,inventar,ram,os,location1,description1,description2,mac,ip,location2,dns,description3
# replace ',' by space
# replace '<>' by space
# replace ';' by ','
# start ps-Script
# import in Zabbix

$hosts = @" 
"@

foreach ($pc in $computers) {

$type = $pc.type
$hwid = $pc.hwid
$serial = $pc.serial
$inv = $pc.inventar
$ram = $pc.ram
$os = $pc.os
$location = $pc.location1
$location2 = $pc.location2
$mac = $pc.mac
$ip = $pc.ip
$dns = $pc.dns

$desc1 = $pc.description1
$desc2 = $pc.description2
$desc3 = $pc.description3
$description = @"
$desc1 $desc2 $desc3
"@





$name = $hwid

Write-Output $description

if (-not ([string]::IsNullOrEmpty($dns)) -and -not ($dns -eq "no-dns") -and -not ($dns -eq "no_dns"))
{
  $name,$reststring = $dns -split "\.",2
}

$hostname = $name

#if (-not ([string]::IsNullOrEmpty($ip)))
#{
#    $hostname = $ip
#}

$singlehost1 = @"
        <host>
            <host>$hostname</host>
            <name>$name</name>
            <groups>
                <group>
                    <name>Imported hosts</name>
                </group>
            </groups>

"@

$singlehost2 = @"
            <tags>
                <tag>
                    <tag>imported</tag>
                    <value>true</value>
                </tag>
            </tags>
            <inventory>
                <type>$type</type>
                <name>$hwid</name>
                <alias>$hostname</alias>
                <os>$os</os>
                <serialno_a>$serial</serialno_a>
                <serialno_b>$inv</serialno_b>
                <macaddress_a>$mac</macaddress_a>
                <location>$location2 $location1</location>
                <notes>$description</notes>
                <host_router>$dns</host_router>
                <oob_ip>$ip</oob_ip>
            </inventory>
        </host>

"@

if (-not ([string]::IsNullOrEmpty($ip)))
{
    $interface = @"
            <interfaces>
                <interface>
                    <type>SNMP</type>
                    <ip>$ip</ip>
                    <port>161</port>
                    <details>
                        <community>{`$SNMP_COMMUNITY}</community>
                    </details>
                    <interface_ref>if1</interface_ref>
                </interface>
            </interfaces>

"@
    $singlehost = $singlehost1 + $interface + $singlehost2
}
else {
    $singlehost = $singlehost1 + $singlehost2
}


$hosts = $hosts + $singlehost
}




$xml = @"
<?xml version="1.0" encoding="UTF-8"?>
<zabbix_export>
    <version>5.2</version>
    <date>2020-11-24T09:27:01Z</date>
    <groups>
        <group>
            <name>Discovered hosts</name>
        </group>
    </groups>
    <hosts>
    $hosts
    </hosts>
</zabbix_export>
"@




$Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $False
[System.IO.File]::WriteAllLines($xmlPath, $xml, $Utf8NoBomEncoding)