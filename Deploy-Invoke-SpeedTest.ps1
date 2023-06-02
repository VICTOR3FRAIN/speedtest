$path_psm = ($env:PSModulePath.Split(";")[0])+"\Invoke-SpeedTest\Invoke-SpeedTest.psm1"
if (!(Test-Path $path_psm)) {
    New-Item $path_psm -ItemType File -Force
}
