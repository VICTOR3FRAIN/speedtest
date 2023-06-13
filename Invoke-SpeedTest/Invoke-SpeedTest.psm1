function Invoke-SpeedTestConserva {
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
Invoke-SpeedTest -LogWrite -LogPath "$home\Documents\Ookla-SpeedTest-Log.txt" # Set default path for log
Invoke-SpeedTest -LogRead | ft # Out log to PSObject
Invoke-SpeedTest -LogClear # Clear log file
#>
param(
    [switch]$LogWrite,
    $LogPath = "$home\Documents\SpeedTest-Log.txt",
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

    [void][System.Reflection.Assembly]::LoadWithPartialName("C:\Program Files (x86)\MySQL\MySQL Connector NET 8.0.332\MySql.Data")
    $MySQLAdminUserName = 'consultores_speedtest'
    $MySQLAdminPassword = 'EfcD2RzPvykzWfbAMiRddAjwQTFRK6MT2sr75fwC...'
    $MySQLDatabase = 'speedtest'
    $MySQLHost = '187.188.165.208'
    $ConnectionString = "server=" + $MySQLHost + ";port=6001;uid=" + $MySQLAdminUserName + ";pwd=" + $MySQLAdminPassword + ";database="+$MySQLDatabase

    ## inicio - inserta en archivo txt
        if ($LogWrite) {
        
        id_cata_sucursal = 32
        connection_icon = $Data.result.connection_icon
        download = $Data.result.download
        upload = $Data.result.upload
        latency = $Data.result.latency
        distance = $Data.result.distance
        server_id = $Data.result.server_id
        sponsor_name = $Data.result.sponsor_name
        isp_name = $Data.result.isp_name
        idle_latency = $Data.result.idle_latency
        download_latency = $Data.result.download_latency
        upload_latency = $Data.result.upload_latency
        fecha_consulta = $Data.result.fecha_consulta

        $Query = "insert into reg_diario (id_cata_sucursal,connection_icon,download,upload,latency,distance,server_id,sponsor_name,isp_name,idle_latency,download_latency,upload_latency,fecha_consulta) VALUE ($id_cata_sucursal,$connection_icon,$download,$upload,$latency,$distance,$server_id,$sponsor_name,$isp_name,$idle_latency,$download_latency,$upload_latency,$fecha_consulta)"
        
        $Connection = New-Object MySql.Data.MySqlClient.MySqlConnection
        $Connection.ConnectionString = $ConnectionString
        $Connection.Open()
            $Command = New-Object MySql.Data.MySqlClient.MySqlCommand($Query, $Connection)
            $DataAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($Command)
            $DataSet = New-Object System.Data.DataSet
            $RecordCount = $dataAdapter.Fill($dataSet, "data")
            $DataSet.Tables[0]
        $Connection.Close()

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
 
    ## fin insert archivo txt
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
