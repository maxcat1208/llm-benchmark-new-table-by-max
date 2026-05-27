$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$runnerPath = Join-Path $repoRoot "scripts\run_full_benchmark.ps1"
$content = Get-Content -LiteralPath $runnerPath -Raw

if ($content -notmatch '\$redacted = \$text') {
    throw "Redact-TextFile should compare redacted content with original content before writing."
}

if ($content -notmatch 'if \(\$redacted -eq \$text\) \{\s*return\s*\}') {
    throw "Redact-TextFile must return without Set-Content when no redaction is needed, otherwise watchdog progress can be faked by redaction writes."
}

"ok"
