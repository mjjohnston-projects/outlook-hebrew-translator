$ErrorActionPreference = 'Stop'

$appRoot = Split-Path -Parent $PSScriptRoot
$node = (Get-Command node -ErrorAction Stop).Source
Set-Location $appRoot

# Run in the foreground so Task Scheduler can restart it if the server exits unexpectedly.
& $node server.mjs
exit $LASTEXITCODE
