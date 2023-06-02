function Invoke-SpeedTest {
<#
.SYNOPSIS
Module creating metrics measurements Internet speed to mode cli (no use dependencies) for output to console PSObject and log file
Data collection resource: speedtest.net (dev Ookla)
Using native API method (via InternetExplorer) for web function start
Using REST API GET method (via Invoke-RestMethod) for parsing JSON report
Works in PSVersion 5.1 (PowerShell 7.3 not supported)
.DESCRIPTION
Example:
$SpeedTest = Invoke-SpeedTest # Output to variable full report
Invoke-SpeedTest -LogWrite # Write to log
Invoke-SpeedTest -LogWrite -LogPath "$home\Documents\Conserva-SpeedTest-Log.txt" # Set default path for log
Invoke-SpeedTest -LogRead | ft # Out log to PSObject
Invoke-SpeedTest -LogClear # Clear log file
.LINK
https://github.com/Lifailon/Ookla-SpeedTest-API
#>
param(
    [switch]$LogWrite,
    $LogPath = "$home\Documents\Conserva-SpeedTest-Log.txt",
    [switch]$LogRead,
    [switch]$LogClear
)

if ($LogClear) {
    $null > $LogPath
    return
}

if (!$LogRead) {
$ie = New-Object -ComObject InternetExplorer.Application
$ie.navigate("https://www.speedtest.net")

while ($True) {
    if ($ie.ReadyState -eq 4) {
        break
    } else {
        sleep -Milliseconds 100
    }
}

$SPAN_Elements = $ie.document.IHTMLDocument3_getElementsByTagName("SPAN")
$Go_Button = $SPAN_Elements | ? innerText -like "go"
$Go_Button.Click()

### Get result URL
$Source_URL = $ie.LocationURL
$Sec = 0
while ($True) {
    if ($ie.LocationURL -notlike $Source_URL) {
        Write-Progress -Activity "SpeedTest Completed" -PercentComplete 100
        $Result_URL = $ie.LocationURL
        $ie.Quit()
        break
    } else {
        sleep 1
        $Sec += 1
        Write-Progress -Activity "Started SpeedTest" -Status "Run time: $Sec sec" -PercentComplete $Sec
    }
}

### Parsing Web Content (JSON)
$Cont = irm $Result_URL
$Data = ($Cont -split "window.OOKLA.")[3] -replace "(.+ = )|(;)" | ConvertFrom-Json

### Convert Unix Time
$EpochTime = [DateTime]"1/1/1970"
$TimeZone = Get-TimeZone
$UTCTime = $EpochTime.AddSeconds($Data.result.date)
$Data.result.date = $UTCTime.AddMinutes($TimeZone.BaseUtcOffset.TotalMinutes)


    $mysql_server = "187.188.165.208:6001"
    $mysql_user = "consultores_speedtest "
    $mysql_password = "EfcD2RzPvykzWfbAMiRddAjwQTFRK6MT2sr75fwC..."
    $dbName = "speedtest"
    [void][system.reflection.Assembly]::LoadFrom("C:\Program Files (x86)\MySQL\MySQL Installer for Windows\MySql.Data.dll")
    $Connection = New-Object -TypeName MySql.Data.MySqlClient.MySqlConnection
    $Connection.ConnectionString = "SERVER=$mysql_server;DATABASE=$dbName;UID=$mysql_user;PWD=$mysql_password"
    $Connection.Open()

    ## inicio - inserta en archivo txt
        if ($LogWrite) {
        $sql = New-Object MySql.Data.MySqlClient.MySqlCommand
        $sql.Connection = $Connection
        $sql.CommandText = "insert into reg_diario (id_cata_sucursal,connection_icon,download,upload,latency,distance,server_id,sponsor_name,isp_name,idle_latency,download_latency,upload_latency,fecha_consulta) VALUE (32,$Data.result.connection_icon,$Data.result.download,$Data.result.upload,$Data.result.latency,$Data.result.distance,$Data.result.server_id,$Data.result.sponsor_name,$Data.result.isp_name,$Data.result.idle_latency,$Data.result.download_latency,$Data.result.upload_latency,$Data.result.fecha_consulta)"
        $sql.ExecuteNonQuery()

$time = $Data.result.date
$ping = $Data.result.idle_latency

$Download = [string]($Data.result.download)
$d2 = $Download[-3..-1] -Join ""
$d1 = $Download[-10..-4] -Join ""
$down = "$d1.$d2 Mbit"

$Upload = [string]($Data.result.upload)
$u2 = $Upload[-3..-1] -Join ""
$u1 = $Upload[-10..-4] -Join ""
$up = "$u1.$u2 Mbit"

$Out_Log = "$time  Download: $down  Upload: $up  Ping latency: $ping ms"
$Out_Log >> $LogPath
}

$Data.result
}

if ($LogRead) {
    $gcLog = gc $LogPath
    $Collections = New-Object System.Collections.Generic.List[System.Object]
    foreach ($gcl in $gcLog) {
        $out = $gcl -split "\s\s"
        $dt  = $out[0] -split "\s"
        $Collections.Add([PSCustomObject]@{
            Date        = $dt[0];
            Time        = $dt[1];
            Download    = $out[1] -replace "Download: ";
            Upload      = $out[2] -replace "Upload: ";
            Ping        = $out[3] -replace "Ping latency: ";
        })
    }
    $Collections
}
}
