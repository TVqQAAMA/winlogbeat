Add-Type -AssemblyName System.IO.Compression.FileSystem
function Unzip
{
    param([string]$zipfile, [string]$outpath)

    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
}

$source = '\\dc.internal.local\sysvol\internal.local\winlogbeat\winlogbeat-7.4.2-windows-x86_64.zip'
$cert = '\\dc.internal.local\sysvol\internal.local\winlogbeat\logstash-forwarder.crt'
$config = '\\dc.internal.local\sysvol\internal.local\winlogbeat\winlogbeat.yml'
$dest = 'C:\Program Files\Winlogbeat'

if (-not (Test-Path $dest))
{    
    New-Item -Path $dest -ItemType "directory"
    Copy-Item $source -Destination $dest
    $filename = Get-ChildItem -Path $dest -Force -Recurse -File | Select-Object -First 1
    Unzip $dest\$filename $dest
    Copy-Item $cert -Destination $dest

    $workdir = $dest + '\' + $filename
    $workdir = $workdir -replace ".{4}$" # remove .zip

    Write-Host $config
    
    Remove-Item $workdir\winlogbeat.yml
    Copy-Item $config -Destination $workdir  

    # Create the new service.
    New-Service -name winlogbeat `
      -displayName Winlogbeat `
      -binaryPathName "`"$workdir\winlogbeat.exe`" -c `"$workdir\winlogbeat.yml`" -path.home `"$workdir`" -path.data `"C:\ProgramData\winlogbeat`" -path.logs `"C:\ProgramData\winlogbeat\logs`" -E logging.files.redirect_stderr=true"

    # Attempt to set the service to delayed start using sc config.

    Start-Process -FilePath sc.exe -ArgumentList 'config winlogbeat start= delayed-auto'
    
    Remove-Item $dest\$filename
}