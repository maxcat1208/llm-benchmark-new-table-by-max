$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$runnerPath = Join-Path $repoRoot "scripts\run_full_benchmark.ps1"
$content = Get-Content -LiteralPath $runnerPath -Raw

if ($content -match '\[int\[\]\]\$ParallelLevels') {
    throw "run_full_benchmark.ps1 must not bind -ParallelLevels as [int[]] directly. External powershell.exe -File calls can turn '8,16,32,64' into 8163264."
}

if ($content -notmatch 'ConvertTo-ParallelLevels') {
    throw "run_full_benchmark.ps1 should normalize -ParallelLevels through ConvertTo-ParallelLevels."
}

"ok"
