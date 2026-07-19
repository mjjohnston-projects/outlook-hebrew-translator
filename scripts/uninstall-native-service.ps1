$ErrorActionPreference = 'Stop'
$serviceName = 'OutlookEmailLanguageAssistant'
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]::new($identity)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  throw 'Administrator rights are required. Run this script from an elevated PowerShell window.'
}

Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
& sc.exe delete $serviceName | Out-Null
Write-Host 'Removed the Outlook Email Language Assistant Windows service.'
