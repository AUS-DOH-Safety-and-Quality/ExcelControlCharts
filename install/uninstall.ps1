# Unregisters the ExcelControlCharts add-in and removes its manifest.
# Run directly, or pipe from the web:
#   irm https://aus-doh-safety-and-quality.github.io/ExcelControlCharts/uninstall.ps1 | iex

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$DataDir      = Join-Path $env:APPDATA 'ExcelControlCharts'
$ManifestPath = Join-Path $DataDir 'manifest.xml'
$DeveloperKey = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer'

if (Test-Path -LiteralPath $ManifestPath) {
    $AddInId = ([xml](Get-Content -Raw -LiteralPath $ManifestPath)).OfficeApp.Id
    if ($AddInId) {
        Remove-ItemProperty -Path $DeveloperKey -Name $AddInId -ErrorAction SilentlyContinue
    }
}

# Older versions of the tooling keyed the value by manifest path.
Remove-ItemProperty -Path $DeveloperKey -Name $ManifestPath -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $DataDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host 'Unregistered the add-in. Restart Excel to finish removing it.'
