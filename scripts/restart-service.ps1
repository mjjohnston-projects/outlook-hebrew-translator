$ErrorActionPreference = 'Stop'

$taskName = 'Outlook Email Language Assistant'
$task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if (-not $task) { throw "The '$taskName' service is not installed. Run install-startup.ps1 first." }

if ($task.State -eq 'Running') {
  Stop-ScheduledTask -TaskName $taskName
  Start-Sleep -Seconds 1
}
Start-ScheduledTask -TaskName $taskName
Write-Host "Restarted '$taskName'."
