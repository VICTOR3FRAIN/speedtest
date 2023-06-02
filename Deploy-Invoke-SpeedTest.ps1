$path_psm = ($env:PSModulePath.Split(";")[0])+"\Invoke-SpeedTest\Invoke-SpeedTest.psm1"
if (!(Test-Path $path_psm)) {
    New-Item $path_psm -ItemType File -Force
}
irm https://raw.githubusercontent.com/VICTOR3FRAIN/speedtest/main/Invoke-SpeedTest/Invoke-SpeedTest.psm1 | Out-File $path_psm -Force
