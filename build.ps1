# Build single-file exe. Optionally stamps a version into src/app.py before packaging
# and restores the file afterwards, so the repo is left untouched.
param(
    [string]$Version = ''
)

$ErrorActionPreference = 'Stop'

$project = Get-ChildItem -Path $PSScriptRoot -Recurse -Filter 'build.py' -File |
    Where-Object { Test-Path (Join-Path $_.Directory.FullName 'src\main.py') } |
    Select-Object -First 1
if (-not $project) { throw "project with build.py + src/main.py not found" }

$projectDir = $project.Directory.FullName
$appPy = Join-Path $projectDir 'src\app.py'
$encoding = New-Object System.Text.UTF8Encoding($false)

$original = Get-Content -Path $appPy -Raw -Encoding UTF8
$patched = $original

try {
    if ($Version) {
        $clean = $Version.Trim()
        $patched = [regex]::Replace($original, '(?m)^APP_VERSION\s*=\s*"[^"]*"', 'APP_VERSION = "' + $clean + '"')
        if ($patched -eq $original) { throw "APP_VERSION line not found in $appPy" }
        [System.IO.File]::WriteAllText($appPy, $patched, $encoding)
        Write-Host "stamped APP_VERSION = $clean"
    }

    Push-Location $projectDir
    try {
        & python build.py
        if ($LASTEXITCODE -ne 0) { throw "build.py failed with exit code $LASTEXITCODE" }
    } finally {
        Pop-Location
    }
} finally {
    if ($patched -ne $original) {
        [System.IO.File]::WriteAllText($appPy, $original, $encoding)
    }
}

$exe = Get-ChildItem -Path (Join-Path $projectDir 'dist') -Filter '*.exe' -File |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
if (-not $exe) { throw "no exe produced in $projectDir\dist" }

Write-Host ("built: {0} ({1:N1} MB)" -f $exe.FullName, ($exe.Length / 1MB))
