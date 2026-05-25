param(
    [Parameter(Mandatory = $true)]
    [string]$RunDir,

    [Parameter(Mandatory = $true)]
    [string]$SourceArgs,

    [string]$TemplateXlsx = "",

    [string]$ExcelModelName = "Kimi2.5",

    [int[]]$ParallelLevels = @(8, 16, 32, 64),

    [int]$OutputTokens = 2048,

    [int]$TotalTimeoutSec = 21600,

    [int]$NoProgressTimeoutSec = 1800,

    [int]$PollIntervalSec = 60,

    [string]$EvalScopeCondaEnv = "evalscope"
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptDir

$RunDir = [System.IO.Path]::GetFullPath($RunDir)
$logsDir = Join-Path $RunDir "logs"
$outputsDir = Join-Path $RunDir "outputs"
$shortRoot = Join-Path $outputsDir "short"
$longRoot = Join-Path $outputsDir "long"
$summaryPath = Join-Path $RunDir "run-summary.md"
$statusPath = Join-Path $RunDir "status.json"
$targetXlsx = Join-Path $RunDir "模型测试汇总表-new.xlsx"

New-Item -ItemType Directory -Force -Path $logsDir, $shortRoot, $longRoot | Out-Null

if ($TotalTimeoutSec -le 0) {
    throw "TotalTimeoutSec must be greater than 0"
}
if ($NoProgressTimeoutSec -le 0) {
    throw "NoProgressTimeoutSec must be greater than 0"
}
if ($PollIntervalSec -le 0) {
    throw "PollIntervalSec must be greater than 0"
}

function Write-RunLog {
    param([string]$Message)
    $line = "$(Get-Date -Format o) $Message"
    Add-Content -LiteralPath (Join-Path $logsDir "run.log") -Value $line -Encoding UTF8
    Write-Output $line
}

function Redact-TextFile {
    param([string]$Path, [string]$Secret)
    if (-not (Test-Path -LiteralPath $Path)) {
        return
    }
    $text = Get-Content -LiteralPath $Path -Raw -ErrorAction SilentlyContinue
    if ($null -eq $text) {
        return
    }
    if ($Secret) {
        $text = $text.Replace($Secret, "[REDACTED]")
    }
    $text = $text -replace "sk-[A-Za-z0-9_-]+", "sk-[REDACTED]"
    Set-Content -LiteralPath $Path -Value $text -NoNewline -Encoding UTF8
}

function Redact-BenchmarkArgs {
    param([string]$Root, [string]$Secret)
    if (-not (Test-Path -LiteralPath $Root)) {
        return
    }
    Get-ChildItem -LiteralPath $Root -Recurse -File -Filter "benchmark_args.json" | ForEach-Object {
        try {
            $raw = Get-Content -LiteralPath $_.FullName -Raw
            $obj = $raw | ConvertFrom-Json
            if ($obj.api_key) {
                $obj.api_key = "[REDACTED]"
            }
            if ($obj.headers -and $obj.headers.Authorization) {
                $obj.headers.Authorization = "Bearer [REDACTED]"
            }
            $obj | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $_.FullName -Encoding UTF8
        } catch {
            Redact-TextFile -Path $_.FullName -Secret $Secret
        }
    }
}

function Redact-OutputTextFiles {
    param([string]$Root, [string]$Secret)
    if (-not (Test-Path -LiteralPath $Root)) {
        return
    }

    Get-ChildItem -LiteralPath $Root -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in @(".json", ".log", ".txt", ".md", ".ps1") } |
        ForEach-Object {
            Redact-TextFile -Path $_.FullName -Secret $Secret
        }
}

function Get-ModelOutputDir {
    param([string]$Root)
    $rootItem = Get-Item -LiteralPath $Root
    $direct = Get-ChildItem -LiteralPath $rootItem.FullName -Directory -Filter "parallel_*_number_*" -ErrorAction SilentlyContinue
    if ($direct.Count -gt 0) {
        return $rootItem.FullName
    }
    $candidates = Get-ChildItem -LiteralPath $rootItem.FullName -Recurse -Directory -Filter "parallel_*_number_*" -ErrorAction SilentlyContinue |
        Group-Object { $_.Parent.FullName } |
        Sort-Object Count -Descending
    if ($candidates.Count -eq 0) {
        throw "Cannot find EvalScope model output under $Root"
    }
    return $candidates[0].Name
}

function ConvertTo-ProcessArgument {
    param([string]$Value)

    if ($null -eq $Value) {
        return '""'
    }
    if ($Value.Length -eq 0) {
        return '""'
    }
    if ($Value -notmatch '[\s"]') {
        return $Value
    }

    $builder = [System.Text.StringBuilder]::new()
    [void]$builder.Append('"')
    $backslashes = 0

    foreach ($ch in $Value.ToCharArray()) {
        if ($ch -eq [char]92) {
            $backslashes += 1
            continue
        }
        if ($ch -eq [char]34) {
            if ($backslashes -gt 0) {
                [void]$builder.Append('\' * ($backslashes * 2))
                $backslashes = 0
            }
            [void]$builder.Append('\"')
            continue
        }
        if ($backslashes -gt 0) {
            [void]$builder.Append('\' * $backslashes)
            $backslashes = 0
        }
        [void]$builder.Append($ch)
    }

    if ($backslashes -gt 0) {
        [void]$builder.Append('\' * ($backslashes * 2))
    }
    [void]$builder.Append('"')
    return $builder.ToString()
}

function Get-OutputProgress {
    param([string]$Root)

    if (-not (Test-Path -LiteralPath $Root)) {
        return [pscustomobject]@{
            FileCount = 0
            TotalBytes = 0
            NewestWriteUtc = $null
            NewestFile = ""
            Signature = "0|0|"
        }
    }

    $files = @(Get-ChildItem -LiteralPath $Root -Recurse -File -Force -ErrorAction SilentlyContinue)
    $totalBytes = [int64]0
    $newestWriteUtc = $null
    $newestFile = ""

    foreach ($file in $files) {
        $totalBytes += [int64]$file.Length
        if ($null -eq $newestWriteUtc -or $file.LastWriteTimeUtc -gt $newestWriteUtc) {
            $newestWriteUtc = $file.LastWriteTimeUtc
            $newestFile = $file.FullName
        }
    }

    $stamp = ""
    if ($null -ne $newestWriteUtc) {
        $stamp = $newestWriteUtc.ToString("o")
    }

    return [pscustomobject]@{
        FileCount = $files.Count
        TotalBytes = $totalBytes
        NewestWriteUtc = $newestWriteUtc
        NewestFile = $newestFile
        Signature = "$($files.Count)|$totalBytes|$stamp"
    }
}

function Get-ChildProcessIds {
    param([int]$ParentProcessId)

    $children = @(Get-CimInstance Win32_Process -Filter "ParentProcessId = $ParentProcessId" -ErrorAction SilentlyContinue)
    foreach ($child in $children) {
        [int]$child.ProcessId
        Get-ChildProcessIds -ParentProcessId ([int]$child.ProcessId)
    }
}

function Stop-ProcessTree {
    param([int]$RootProcessId)

    $ids = @(Get-ChildProcessIds -ParentProcessId $RootProcessId)
    $ids += $RootProcessId
    foreach ($id in ($ids | Where-Object { $_ -and $_ -ne $PID } | Select-Object -Unique)) {
        try {
            Stop-Process -Id $id -Force -ErrorAction Stop
        } catch {
            # The process may have already exited while the watchdog was firing.
        }
    }
}

function Test-EvalScopeSuccessOutput {
    param(
        [string]$Stage,
        [string]$OutputRoot
    )

    try {
        $modelOutputDir = Get-ModelOutputDir -Root $OutputRoot
    } catch {
        return $false
    }

    $checkLog = Join-Path $logsDir "$Stage-success-after-exit.json"
    python (Join-Path $scriptDir "assert_evalscope_success.py") --source-dir $modelOutputDir *> $checkLog
    return $LASTEXITCODE -eq 0
}

function Assert-EvalScopeEnvironment {
    param([string]$CondaEnv)

    $checkLog = Join-Path $logsDir "preflight-evalscope.log"
    $conda = Get-Command conda -ErrorAction SilentlyContinue
    if (-not $conda) {
        throw "Cannot find conda. Install or activate an existing EvalScope environment manually, then rerun. This runner never creates or installs conda environments automatically."
    }

    $checkArgs = @("run", "-n", $CondaEnv, "evalscope", "--help")
    $psi = [System.Diagnostics.ProcessStartInfo]::new()
    $psi.FileName = "conda"
    $psi.Arguments = (($checkArgs | ForEach-Object { ConvertTo-ProcessArgument -Value ([string]$_) }) -join " ")
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true

    $process = [System.Diagnostics.Process]::new()
    $process.StartInfo = $psi
    try {
        [void]$process.Start()
        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        $exitCode = $process.ExitCode
    } finally {
        if ($process) {
            $process.Dispose()
        }
    }

    Set-Content -LiteralPath $checkLog -Value ($stdout + $stderr) -Encoding UTF8
    if ($exitCode -ne 0) {
        throw "Cannot run 'conda run -n $CondaEnv evalscope --help'. Use an existing EvalScope conda environment or pass -EvalScopeCondaEnv with the correct environment name. Do not auto-install from this skill. See $checkLog"
    }
}

function Invoke-EvalScopePerf {
    param(
        [string]$Stage,
        [string]$Dataset,
        [string]$OutputRoot,
        [string]$LogPath,
        [string]$ApiUrl,
        [string]$ApiKey,
        [string]$ModelName,
        [string]$EvalScopeCondaEnv
    )

    $childScriptPath = Join-Path $logsDir "$Stage-evalscope-child.ps1"
    $childScript = @'
param(
    [Parameter(Mandatory = $true)]
    [string]$Dataset,

    [Parameter(Mandatory = $true)]
    [string]$OutputRoot,

    [Parameter(Mandatory = $true)]
    [string]$LogPath,

    [Parameter(Mandatory = $true)]
    [string]$ApiUrl,

    [Parameter(Mandatory = $true)]
    [string]$ModelName,

    [Parameter(Mandatory = $true)]
    [string]$ParallelLevelsCsv,

        [Parameter(Mandatory = $true)]
        [int]$OutputTokens,

        [Parameter(Mandatory = $true)]
        [int]$TotalTimeoutSec,

        [Parameter(Mandatory = $true)]
        [string]$EvalScopeCondaEnv
)

$ErrorActionPreference = "Stop"

try {
$parallelLevels = @($ParallelLevelsCsv -split "," | ForEach-Object { [int]$_ })
$numbers = @()
foreach ($p in $parallelLevels) {
    $numbers += ($p * 10)
}

$apiKey = $env:LLM_BENCH_API_KEY
if (-not $apiKey) {
    throw "LLM_BENCH_API_KEY is not set"
}

$cmdArgs = @(
    "run", "-n", $EvalScopeCondaEnv, "evalscope", "perf",
    "--url", $ApiUrl,
    "--api-key", $apiKey,
    "--model", $ModelName,
    "--dataset", $Dataset,
    "--min-tokens", [string]$OutputTokens,
    "--max-tokens", [string]$OutputTokens,
    "--api", "openai",
    "--parallel"
)
$cmdArgs += ($parallelLevels | ForEach-Object { [string]$_ })
$cmdArgs += "--number"
$cmdArgs += ($numbers | ForEach-Object { [string]$_ })
$cmdArgs += @(
    "--total-timeout", [string]$TotalTimeoutSec,
    "--extra-args", '{""use_cache"": false, ""ignore_eos"": true}',
    "--outputs-dir", $OutputRoot,
    "--no-timestamp"
)

$oldErrorActionPreference = $ErrorActionPreference
$ErrorActionPreference = "Continue"
& conda @cmdArgs *> $LogPath
$exitCode = $LASTEXITCODE
$ErrorActionPreference = $oldErrorActionPreference
exit $exitCode
} catch {
    $_ | Out-String | Add-Content -LiteralPath $LogPath -Encoding UTF8
    exit 1
}
'@
    Set-Content -LiteralPath $childScriptPath -Value $childScript -Encoding UTF8

    $processArgs = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", $childScriptPath,
        "-Dataset", $Dataset,
        "-OutputRoot", $OutputRoot,
        "-LogPath", $LogPath,
        "-ApiUrl", $ApiUrl,
        "-ModelName", $ModelName,
        "-ParallelLevelsCsv", ($ParallelLevels -join ","),
        "-OutputTokens", [string]$OutputTokens,
        "-TotalTimeoutSec", [string]$TotalTimeoutSec,
        "-EvalScopeCondaEnv", $EvalScopeCondaEnv
    )

    $psi = [System.Diagnostics.ProcessStartInfo]::new()
    $psi.FileName = "powershell.exe"
    $psi.Arguments = (($processArgs | ForEach-Object { ConvertTo-ProcessArgument -Value ([string]$_) }) -join " ")
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.EnvironmentVariables["LLM_BENCH_API_KEY"] = $ApiKey

    $process = $null
    $exitCode = $null
    $watchdogMessage = $null
    $startedAt = Get-Date
    $lastProgressAt = $startedAt
    $progress = Get-OutputProgress -Root $OutputRoot
    $lastProgressSignature = $progress.Signature

    Write-RunLog "START $Stage $Dataset"
    Write-Status -Stage $Stage -State "running" -Details @{
        dataset = $Dataset
        started_at = $startedAt.ToString("o")
        last_progress_at = $lastProgressAt.ToString("o")
        total_timeout_sec = $TotalTimeoutSec
        no_progress_timeout_sec = $NoProgressTimeoutSec
        poll_interval_sec = $PollIntervalSec
    }

    try {
        $process = [System.Diagnostics.Process]::new()
        $process.StartInfo = $psi
        [void]$process.Start()

        Write-RunLog "PID $($process.Id) $Stage $Dataset"

        while (-not $process.WaitForExit($PollIntervalSec * 1000)) {
            $now = Get-Date
            $progress = Get-OutputProgress -Root $OutputRoot
            if ($progress.Signature -ne $lastProgressSignature) {
                $lastProgressSignature = $progress.Signature
                $lastProgressAt = $now
            }

            $elapsedSec = [int]($now - $startedAt).TotalSeconds
            $idleSec = [int]($now - $lastProgressAt).TotalSeconds
            Write-Status -Stage $Stage -State "running" -Details @{
                dataset = $Dataset
                pid = $process.Id
                elapsed_sec = $elapsedSec
                idle_sec = $idleSec
                last_progress_at = $lastProgressAt.ToString("o")
                output_files = $progress.FileCount
                output_bytes = $progress.TotalBytes
                newest_output_file = $progress.NewestFile
                total_timeout_sec = $TotalTimeoutSec
                no_progress_timeout_sec = $NoProgressTimeoutSec
                poll_interval_sec = $PollIntervalSec
            }
            Redact-TextFile -Path $LogPath -Secret $ApiKey
            Redact-OutputTextFiles -Root $OutputRoot -Secret $ApiKey

            if ($elapsedSec -ge $TotalTimeoutSec) {
                $watchdogMessage = "$Stage $Dataset exceeded total timeout ${TotalTimeoutSec}s"
                break
            }
            if ($idleSec -ge $NoProgressTimeoutSec) {
                $watchdogMessage = "$Stage $Dataset had no output progress for ${NoProgressTimeoutSec}s"
                break
            }
        }

        if ($watchdogMessage) {
            Write-RunLog "WATCHDOG $watchdogMessage"
            Write-Status -Stage $Stage -State "failed" -Details @{
                dataset = $Dataset
                pid = $process.Id
                reason = $watchdogMessage
                last_progress_at = $lastProgressAt.ToString("o")
                output_files = $progress.FileCount
                output_bytes = $progress.TotalBytes
                newest_output_file = $progress.NewestFile
            }
            Stop-ProcessTree -RootProcessId $process.Id
            if (-not $process.WaitForExit(15000)) {
                throw "$watchdogMessage; failed to stop process $($process.Id)"
            }
        } else {
            $process.WaitForExit()
        }

        $exitCode = $process.ExitCode
    } finally {
        if ($process) {
            $process.Dispose()
        }
    }

    Redact-TextFile -Path $LogPath -Secret $ApiKey
    Redact-BenchmarkArgs -Root $OutputRoot -Secret $ApiKey
    Redact-OutputTextFiles -Root $OutputRoot -Secret $ApiKey
    if ($watchdogMessage) {
        throw "$watchdogMessage. See $LogPath"
    }
    if ($exitCode -ne 0) {
        if (Test-EvalScopeSuccessOutput -Stage $Stage -OutputRoot $OutputRoot) {
            Write-RunLog "WARN $Stage $Dataset exited with code $exitCode but produced successful requests; continuing"
            Write-Status -Stage $Stage -State "completed_with_warnings" -Details @{
                dataset = $Dataset
                exit_code = $exitCode
                warning = "EvalScope exited non-zero but produced successful requests"
            }
        } else {
            throw "EvalScope $Dataset failed with exit code $exitCode. See $LogPath"
        }
    }
    Write-RunLog "DONE $Stage $Dataset"
}

function Write-Status {
    param(
        [string]$Stage,
        [string]$State,
        [hashtable]$Details = @{}
    )
    $status = [ordered]@{
        stage = $Stage
        state = $State
        updated_at = (Get-Date -Format o)
        run_dir = $RunDir
    }
    foreach ($key in $Details.Keys) {
        $status[$key] = $Details[$key]
    }
    $status | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $statusPath -Encoding UTF8
}

try {
    Write-Status -Stage "prepare" -State "running"
    Write-RunLog "Run directory: $RunDir"
    Write-RunLog "Checking existing EvalScope environment: $EvalScopeCondaEnv"
    Assert-EvalScopeEnvironment -CondaEnv $EvalScopeCondaEnv

    $source = Get-Content -LiteralPath $SourceArgs -Raw | ConvertFrom-Json
    $apiUrl = $source.url
    $apiKey = $source.api_key
    if (-not $apiKey -and $source.headers -and $source.headers.Authorization) {
        $apiKey = [string]$source.headers.Authorization
        $apiKey = $apiKey -replace "^Bearer\s+", ""
    }
    $modelName = $source.model
    if (-not $apiUrl -or -not $apiKey -or -not $modelName) {
        throw "Source benchmark args must contain url, api_key, and model"
    }

    [ordered]@{
        model = $modelName
        excel_model_name = $ExcelModelName
        api_url = $apiUrl
        api_key = "[REDACTED]"
        parallel_levels = $ParallelLevels
        numbers = ($ParallelLevels | ForEach-Object { $_ * 10 })
        output_tokens = $OutputTokens
        total_timeout_sec = $TotalTimeoutSec
        no_progress_timeout_sec = $NoProgressTimeoutSec
        poll_interval_sec = $PollIntervalSec
        evalscope_conda_env = $EvalScopeCondaEnv
        run_dir = $RunDir
        template_xlsx = $TemplateXlsx
        short_dataset = "speed_benchmark"
        long_dataset = "speed_benchmark_long"
    } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $RunDir "run-config.redacted.json") -Encoding UTF8

    if (-not $TemplateXlsx) {
        $bundledTemplate = Join-Path $repoRoot "templates\模型测试汇总表-new.xlsx"
        if (Test-Path -LiteralPath $bundledTemplate) {
            $TemplateXlsx = $bundledTemplate
        }
    }
    if ($TemplateXlsx -and (Test-Path -LiteralPath $TemplateXlsx)) {
        Copy-Item -LiteralPath $TemplateXlsx -Destination $targetXlsx -Force
    }

    Write-Status -Stage "short" -State "running"
    Invoke-EvalScopePerf -Stage "short" -Dataset "speed_benchmark" -OutputRoot $shortRoot -LogPath (Join-Path $logsDir "short.log") -ApiUrl $apiUrl -ApiKey $apiKey -ModelName $modelName -EvalScopeCondaEnv $EvalScopeCondaEnv
    $shortModelDir = Get-ModelOutputDir -Root $shortRoot
    python (Join-Path $scriptDir "assert_evalscope_success.py") --source-dir $shortModelDir *> (Join-Path $logsDir "short-success.json")
    if ($LASTEXITCODE -ne 0) {
        throw "Short benchmark produced no successful requests. See $(Join-Path $logsDir 'short-success.json') and $(Join-Path $logsDir 'short.log')"
    }
    python (Join-Path $scriptDir "fill_new_benchmark_excel.py") --source-dir $shortModelDir --target-xlsx $targetXlsx --sheet-name "speed_benchmark" --excel-model-name $ExcelModelName --run-model-name $modelName *> (Join-Path $logsDir "fill-short.log")
    if ($LASTEXITCODE -ne 0) {
        throw "Filling short sheet failed. See $(Join-Path $logsDir 'fill-short.log')"
    }

    Write-Status -Stage "long" -State "running"
    Invoke-EvalScopePerf -Stage "long" -Dataset "speed_benchmark_long" -OutputRoot $longRoot -LogPath (Join-Path $logsDir "long.log") -ApiUrl $apiUrl -ApiKey $apiKey -ModelName $modelName -EvalScopeCondaEnv $EvalScopeCondaEnv
    $longModelDir = Get-ModelOutputDir -Root $longRoot
    python (Join-Path $scriptDir "assert_evalscope_success.py") --source-dir $longModelDir *> (Join-Path $logsDir "long-success.json")
    if ($LASTEXITCODE -ne 0) {
        throw "Long benchmark produced no successful requests. See $(Join-Path $logsDir 'long-success.json') and $(Join-Path $logsDir 'long.log')"
    }
    python (Join-Path $scriptDir "fill_new_benchmark_excel.py") --source-dir $longModelDir --target-xlsx $targetXlsx --sheet-name "speed_benchmark -long" --excel-model-name $ExcelModelName --run-model-name $modelName *> (Join-Path $logsDir "fill-long.log")
    if ($LASTEXITCODE -ne 0) {
        throw "Filling long sheet failed. See $(Join-Path $logsDir 'fill-long.log')"
    }

    $summaryLines = @(
        "# LLM Benchmark Run Summary",
        "",
        "- Run dir: $RunDir",
        "- Model: $modelName",
        "- Excel model name: $ExcelModelName",
        "- Short output: $shortModelDir",
        "- Long output: $longModelDir",
        "- Workbook: $targetXlsx",
        "- Status: completed"
    )
    Set-Content -LiteralPath $summaryPath -Value $summaryLines -Encoding UTF8

    Write-Status -Stage "complete" -State "completed"
    Write-RunLog "COMPLETED"
} catch {
    $message = $_.Exception.Message
    if ($apiKey) {
        $message = $message.Replace($apiKey, "[REDACTED]")
    }
    Write-Status -Stage "failed" -State "failed" -Details @{ reason = $message }
    Add-Content -LiteralPath (Join-Path $logsDir "run.log") -Value "$(Get-Date -Format o) FAILED $message" -Encoding UTF8
    $summaryLines = @(
        "# LLM Benchmark Run Summary",
        "",
        "- Run dir: $RunDir",
        "- Status: failed",
        "- Error: $message"
    )
    Set-Content -LiteralPath $summaryPath -Value $summaryLines -Encoding UTF8
    throw
}
