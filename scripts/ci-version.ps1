# Version scheme: MAJOR.MINOR.PATCH
# PATCH goes 0..9, then carries into MINOR; MINOR goes 0..9, then carries into MAJOR.
#   1.2.1 -> 1.2.2 ... 1.2.9 -> 1.3.0 ... 1.9.9 -> 2.0.0
# Release version = max(APP_VERSION, next(highest git tag)).
# With no tags yet, APP_VERSION itself is used for the first release.
param(
    [string]$BaseVersion = '',
    [string]$NextOf = '',
    [switch]$IgnoreTags
)

$ErrorActionPreference = 'Stop'

function Get-NextVersion([string]$value) {
    $value = $value.Trim().TrimStart([char[]]@('v', 'V'))
    if ($value -notmatch '^(\d+)\.(\d+)\.(\d+)$') { throw "invalid version: $value" }
    $major = [int]$Matches[1]
    $minor = [int]$Matches[2]
    $patch = [int]$Matches[3]
    if ($patch -lt 9) {
        $patch = $patch + 1
    } else {
        $patch = 0
        if ($minor -lt 9) {
            $minor = $minor + 1
        } else {
            $minor = 0
            $major = $major + 1
        }
    }
    return "$major.$minor.$patch"
}

if ($NextOf) {
    Write-Output (Get-NextVersion $NextOf)
    return
}

$repoRoot = Split-Path -Parent $PSScriptRoot

if ($BaseVersion.Trim()) {
    $base = $BaseVersion.Trim().TrimStart([char[]]@('v', 'V'))
} else {
    $appPy = Get-ChildItem -Path $repoRoot -Recurse -Filter 'app.py' -File |
        Where-Object { $_.Directory.Name -eq 'src' } |
        Select-Object -First 1
    if (-not $appPy) { throw 'src/app.py not found' }
    $m = [regex]::Match((Get-Content -Path $appPy.FullName -Raw -Encoding UTF8), '(?m)^APP_VERSION\s*=\s*"([^"]+)"')
    if (-not $m.Success) { throw "APP_VERSION not found in $($appPy.FullName)" }
    $base = $m.Groups[1].Value
}
if ($base -notmatch '^\d+\.\d+\.\d+$') { throw "invalid base version: $base" }

$highestTag = $null
if (-not $IgnoreTags) {
    $tags = & git -C $repoRoot tag --list 'v*' 2>$null
    foreach ($tag in $tags) {
        if (-not $tag) { continue }
        $parsed = $tag.Trim().TrimStart([char[]]@('v', 'V'))
        if ($parsed -notmatch '^\d+\.\d+\.\d+$') { continue }
        if (-not $highestTag -or ([version]$parsed).CompareTo([version]$highestTag) -gt 0) {
            $highestTag = $parsed
        }
    }
}

if (-not $highestTag) {
    Write-Output $base
    return
}

$candidate = Get-NextVersion $highestTag
if (([version]$base).CompareTo([version]$candidate) -gt 0) {
    Write-Output $base
} else {
    Write-Output $candidate
}
