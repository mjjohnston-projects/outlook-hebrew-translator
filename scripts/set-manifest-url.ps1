param(
  [Parameter(Mandatory = $true)]
  [ValidatePattern('^https://[^/]+$')]
  [string]$Origin,
  [switch]$ReplaceExisting
)

$ErrorActionPreference = 'Stop'
$manifest = Join-Path (Split-Path -Parent $PSScriptRoot) 'manifest.xml'
$origin = $Origin.TrimEnd('/')
$content = Get-Content -Raw -Path $manifest

if ($content -notmatch 'https://YOUR-HTTPS-DOMAIN') {
  if (-not $ReplaceExisting) {
    throw 'The manifest already has a deployment URL. Run again with -ReplaceExisting to change it deliberately.'
  }
  $content = $content -replace 'https://[^/"< ]+\.ngrok[^/"< ]*', $origin
} else {
  $content = $content.Replace('https://YOUR-HTTPS-DOMAIN', $origin)
}
Set-Content -Path $manifest -Value $content -Encoding utf8
Write-Host "Updated manifest.xml to use $origin"
