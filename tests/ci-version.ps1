# Checks for scripts/ci-version.ps1 (version carry rules)
$ErrorActionPreference = 'Stop'

$script = Join-Path $PSScriptRoot '..\scripts\ci-version.ps1'

$cases = [ordered]@{
    '1.2.0' = '1.2.1'
    '1.2.1' = '1.2.2'
    '1.2.9' = '1.3.0'
    '1.3.9' = '1.4.0'
    '1.9.9' = '2.0.0'
    '0.0.9' = '0.1.0'
    'v1.2.0' = '1.2.1'
}

foreach ($k in $cases.Keys) {
    $got = & $script -NextOf $k
    if ($got -ne $cases[$k]) { throw "$k => $got, expected $($cases[$k])" }
}

$base = & $script -BaseVersion '1.2.1' -IgnoreTags
if ($base -ne '1.2.1') { throw "first release should equal APP_VERSION, got $base" }

Write-Host "ci-version ok ($($cases.Count) carry cases)"
