$ErrorActionPreference = 'Stop'

$appRoot = Split-Path -Parent $PSScriptRoot
$node = (Get-Command node -ErrorAction Stop).Source
$existing = Get-CimInstance Win32_Process -Filter "Name = 'node.exe'" -ErrorAction SilentlyContinue |
  Where-Object { $_.CommandLine -like '*outlook-email-language-assistant*server.mjs*' }

if ($existing) { exit 0 }

$logs = Join-Path $appRoot 'logs'
New-Item -ItemType Directory -Force -Path $logs | Out-Null
$timestamp = Get-Date -Format 'yyyy-MM-dd'

Start-Process -FilePath $node `
  -ArgumentList 'server.mjs' `
  -WorkingDirectory $appRoot `
  -WindowStyle Hidden `
  -RedirectStandardOutput (Join-Path $logs "server-$timestamp.log") `
  -RedirectStandardError (Join-Path $logs "server-$timestamp.error.log")
