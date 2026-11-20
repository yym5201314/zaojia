# Sanity checks for scripts/ci-version.ps1
$ErrorActionPreference = 'Stop'

$script = Join-Path $PSScriptRoot '..\scripts\ci-version.ps1'

$v = & $script -RunNumber 7
if ($v -notmatch '^\d+(\.\d+){1,}\.\d+$') { throw "unexpected version: $v" }
if (-not $v.EndsWith('.7')) { throw "run number not appended: $v" }

$v2 = & $script -RunNumber 12 -BaseVersion 'v2.5'
if ($v2 -ne '2.5.12') { throw "override failed: $v2" }

Write-Host "ci-version ok: $v / $v2"
