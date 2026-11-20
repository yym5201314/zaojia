# Compute build version: <base>.<run_number>
# Base version comes from -BaseVersion, otherwise from APP_VERSION in src/app.py.
param(
    [int]$RunNumber = 1,
    [string]$BaseVersion = ''
)

$ErrorActionPreference = 'Stop'

if ($RunNumber -lt 1) { throw "RunNumber must be >= 1" }

$repoRoot = Split-Path -Parent $PSScriptRoot
$appPy = Get-ChildItem -Path $repoRoot -Recurse -Filter 'app.py' -File |
    Where-Object { $_.Directory.Name -eq 'src' } |
    Select-Object -First 1
if (-not $appPy) { throw "src/app.py not found" }

$base = $BaseVersion.Trim().TrimStart([char[]]@('v', 'V'))
if (-not $base) {
    $text = Get-Content -Path $appPy.FullName -Raw -Encoding UTF8
    $m = [regex]::Match($text, '(?m)^APP_VERSION\s*=\s*"([^"]+)"')
    if (-not $m.Success) { throw "APP_VERSION not found in $($appPy.FullName)" }
    $base = $m.Groups[1].Value
}

if ($base -notmatch '^\d+(\.\d+)*$') { throw "invalid base version: $base" }

Write-Output ("{0}.{1}" -f $base, $RunNumber)
