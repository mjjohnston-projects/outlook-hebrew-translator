$ErrorActionPreference = 'Stop'

$taskName = 'Outlook Email Language Assistant'
$legacyTaskName = 'Outlook Hebrew Translator'
$runner = Join-Path $PSScriptRoot 'run-service.ps1'
$powershell = (Get-Command powershell.exe -ErrorAction Stop).Source

if (-not (Test-Path (Join-Path (Split-Path -Parent $PSScriptRoot) '.env'))) {
  throw 'Create and configure the .env file before enabling startup.'
}

$action = New-ScheduledTaskAction -Execute $powershell -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$runner`""
$trigger = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERDOMAIN\$env:USERNAME"
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Description 'Starts the Outlook Email Language Assistant server when this user signs in.' -Force | Out-Null
Unregister-ScheduledTask -TaskName $legacyTaskName -Confirm:$false -ErrorAction SilentlyContinue
Start-ScheduledTask -TaskName $taskName
Write-Host "Installed and started '$taskName'."
Write-Host "Restart it anytime with: powershell -NoProfile -ExecutionPolicy Bypass -File `"$PSScriptRoot\restart-service.ps1`""
