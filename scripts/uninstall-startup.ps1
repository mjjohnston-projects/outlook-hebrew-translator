$taskName = 'Outlook Hebrew Translator'
Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
Write-Host "Removed '$taskName'."
