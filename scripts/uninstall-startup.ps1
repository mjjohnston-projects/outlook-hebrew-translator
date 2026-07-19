$taskNames = 'Outlook Email Language Assistant', 'Outlook Hebrew Translator'
foreach ($taskName in $taskNames) { Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue }
Write-Host 'Removed the Email Language Assistant startup task.'
