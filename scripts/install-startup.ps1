$ErrorActionPreference = 'Stop'

$taskName = 'Outlook Hebrew Translator'
$runner = Join-Path $PSScriptRoot 'run-translator.ps1'
$powershell = (Get-Command powershell.exe -ErrorAction Stop).Source

if (-not (Test-Path (Join-Path (Split-Path -Parent $PSScriptRoot) '.env'))) {
  throw 'Create and configure the .env file before enabling startup.'
}

$action = New-ScheduledTaskAction -Execute $powershell -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$runner`""
$trigger = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERDOMAIN\$env:USERNAME"
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Description 'Starts the Outlook Hebrew Translator server when this user signs in.' -Force | Out-Null
Start-ScheduledTask -TaskName $taskName
Write-Host "Installed and started '$taskName'."
Write-Host 'Logs are stored in the app\logs folder.'
