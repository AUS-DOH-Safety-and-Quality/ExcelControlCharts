# Registers the deployed ExcelControlCharts add-in for sideloading in Excel.
# Run directly, or pipe from the web:
#   irm https://aus-doh-safety-and-quality.github.io/ExcelControlCharts/install.ps1 | iex

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$BaseUrl = 'https://aus-doh-safety-and-quality.github.io/ExcelControlCharts'

$DataDir      = Join-Path $env:APPDATA 'ExcelControlCharts'
$ManifestPath = Join-Path $DataDir 'manifest.xml'
$DeveloperKey = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer'

New-Item -ItemType Directory -Path $DataDir -Force | Out-Null

# The published manifest already points at the deployment, so no rewriting is needed.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Invoke-WebRequest -Uri "$BaseUrl/manifest.xml" -OutFile $ManifestPath -UseBasicParsing

$AddInId = ([xml](Get-Content -Raw -LiteralPath $ManifestPath)).OfficeApp.Id
if (-not $AddInId) { throw "Could not read <Id> from $ManifestPath." }

New-Item -Path $DeveloperKey -Force | Out-Null

# Drop any previous registration, including values keyed by path rather than id.
Remove-ItemProperty -Path $DeveloperKey -Name $AddInId      -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $DeveloperKey -Name $ManifestPath -ErrorAction SilentlyContinue

New-ItemProperty -Path $DeveloperKey -Name $AddInId -Value $ManifestPath `
                 -PropertyType String -Force | Out-Null

Write-Host "Registered the add-in from $BaseUrl/"
Write-Host 'Restart Excel and choose it from the Home tab.'
