$ErrorActionPreference = 'Stop'

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]::new($identity)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  throw 'Administrator rights are required. Run this script from an elevated PowerShell window.'
}

$serviceName = 'OutlookEmailLanguageAssistant'
$displayName = 'Outlook Email Language Assistant'
$appRoot = Split-Path -Parent $PSScriptRoot
$nodePath = (Get-Command node -ErrorAction Stop).Source
$certificatePath = Join-Path $env:USERPROFILE '.office-addin-dev-certs\localhost.crt'
$keyPath = Join-Path $env:USERPROFILE '.office-addin-dev-certs\localhost.key'
$serviceExe = Join-Path $appRoot 'service\publish\EmailLanguageAssistantService.exe'

foreach ($path in $certificatePath, $keyPath, $serviceExe, (Join-Path $appRoot '.env')) {
  if (-not (Test-Path -LiteralPath $path)) { throw "Required file is missing: $path" }
}

Stop-ScheduledTask -TaskName 'Outlook Email Language Assistant' -ErrorAction SilentlyContinue
Unregister-ScheduledTask -TaskName 'Outlook Email Language Assistant' -Confirm:$false -ErrorAction SilentlyContinue

if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue) {
  Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
  & sc.exe delete $serviceName | Out-Null
  Start-Sleep -Seconds 1
}

$arguments = "--app-root `"$appRoot`" --node-path `"$nodePath`" --certificate `"$certificatePath`" --key `"$keyPath`""
$binaryPath = "`"$serviceExe`" $arguments"
New-Service -Name $serviceName -BinaryPathName $binaryPath -DisplayName $displayName -StartupType Automatic -Description 'Runs the local Outlook Email Language Assistant server.'
& sc.exe description $serviceName 'Runs the local Outlook Email Language Assistant server.'
if ($LASTEXITCODE -ne 0) { throw "Windows could not set the service description (sc.exe exit code $LASTEXITCODE)." }
& sc.exe failure $serviceName reset= 86400 actions= restart/5000/restart/5000/restart/5000
if ($LASTEXITCODE -ne 0) { throw "Windows could not configure service recovery (sc.exe exit code $LASTEXITCODE)." }
Start-Service -Name $serviceName
Write-Host "Installed '$displayName'. Open services.msc to view or manage it."
