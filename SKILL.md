---
name: llm-benchmark-new-table-by-max
description: Use when running EvalScope LLM performance benchmarks or filling the two-sheet 模型测试汇总表-new.xlsx template with percentile and summary metrics from speed_benchmark and speed_benchmark_long outputs.
---

# LLM Benchmark New Table by Max

## 核心原则

这个 skill 使用新版双 sheet Excel 模板：

```text
templates/
  模型测试汇总表-new.xlsx
```

模板有两个工作表：

- `speed_benchmark`：短文本数据集。
- `speed_benchmark -long`：长文本数据集。

每个模型在每个 sheet 里按 `8 / 16 / 32 / 64` 并发分成 4 组，每组 9 行分位数：`1% / 5% / 10% / 25% / 50% / 75% / 90% / 95% / 99%`。

## 环境预检

不要自动安装 EvalScope、conda、Python 包或新环境。先检查用户机器是否已有可用的 EvalScope conda 环境：

```powershell
conda env list
conda run -n evalscope evalscope --help
```

默认环境名是 `evalscope`。如果用户已有环境但名字不同，运行 runner 时传 `-EvalScopeCondaEnv "环境名"`。

如果检查失败，停止并提示用户手动准备环境。不要执行 `pip install`、`conda create`、`conda install` 或类似自动安装命令。

## 完整压测

从已有 EvalScope `benchmark_args.json` 读取 `url`、`api_key`、`model`，按顺序运行短文本、长文本，并分别填入两个 sheet：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File ".\llm-benchmark-new-table-by-max\scripts\run_full_benchmark.ps1" `
  -RunDir ".\llm-benchmark-runs\20260525-Qwen3.5-122B-A10B" `
  -SourceArgs ".\example-output\Qwen3.5-122B-A10B\parallel_8_number_80\benchmark_args.json" `
  -TemplateXlsx ".\llm-benchmark-new-table-by-max\templates\模型测试汇总表-new.xlsx" `
  -ExcelModelName "Qwen3.5-122B-A10B" `
  -EvalScopeCondaEnv "evalscope"
```

如果不传 `TemplateXlsx`，runner 会优先使用包内模板。

runner 会在压测前执行 `conda run -n <EvalScopeCondaEnv> evalscope --help`。环境不存在或 EvalScope 不可用时会直接失败，不会自动安装。

runner 内置 watchdog：默认每 60 秒检查输出目录；如果连续 1800 秒没有输出进展，会停止 EvalScope/conda 子进程树并写入 `status.json`。

EvalScope 在部分请求失败时可能返回非 0 exit code。只要该阶段已经有成功请求，runner 会写 WARN 并继续后续阶段；没有成功请求才停止。

## 只填已有结果

短文本结果填入 `speed_benchmark`：

```powershell
python ".\llm-benchmark-new-table-by-max\scripts\fill_new_benchmark_excel.py" `
  --source-dir ".\outputs\short\Qwen3.5-122B-A10B" `
  --target-xlsx ".\模型测试汇总表-new.xlsx" `
  --sheet-name "speed_benchmark" `
  --excel-model-name "Qwen3.5-122B-A10B" `
  --run-model-name "Qwen3.5-122B-A10B"
```

长文本结果填入 `speed_benchmark -long`：

```powershell
python ".\llm-benchmark-new-table-by-max\scripts\fill_new_benchmark_excel.py" `
  --source-dir ".\outputs\long\Qwen3.5-122B-A10B" `
  --target-xlsx ".\模型测试汇总表-new.xlsx" `
  --sheet-name "speed_benchmark -long" `
  --excel-model-name "Qwen3.5-122B-A10B" `
  --run-model-name "Qwen3.5-122B-A10B"
```

如果 Excel 里找不到 `excel_model_name`，脚本会在对应 sheet 末尾追加一个新模型块。

## Excel 映射

分位数列读取每个并发目录的 `benchmark_percentile.json`：

| Excel 列 | JSON key |
|---|---|
| D `Latency (s)` | `Latency (s)` |
| E `TTFT (ms)` | `TTFT (ms)` |
| F `ITL (ms)` | `ITL (ms)` |
| G `TPOT (ms)` | `TPOT (ms)` |
| H `Input tokens` | `Input tokens` |
| I `Output tokens` | `Output tokens` |
| J `Output (tok/s)` | `Output (tok/s)` |
| K `Total (tok/s)` | `Total (tok/s)` |
| L `Decode (tok/s)` | `Decode (tok/s)` |

汇总列读取每个并发目录的 `benchmark_summary.json`：

| Excel 列 | JSON key |
|---|---|
| M `Test Duration (s)` | `Test Duration (s)` |
| N `Concurrency` | `Concurrency` |
| O `Request Rate (req/s)` | `Request Rate (req/s)` |
| P `Total Requests` | `Total Requests` |
| Q `Success Requests` | `Success Requests` |
| R `Failed Requests` | `Failed Requests` |
| S `Req Throughput (req/s)` | `Req Throughput (req/s)` |
| T `Avg Latency (s)` | `Avg Latency (s)` |
| U `Avg Input Tokens` | `Avg Input Tokens` |
| V `Output Throughput (tok/s)` | `Output Throughput (tok/s)` |
| W `Total Throughput (tok/s)` | `Total Throughput (tok/s)` |
| X `TTFT (ms)` | `TTFT (ms)` |
| Y `TPOT (ms)` | `TPOT (ms)` |
| Z `ITL (ms)` | `ITL (ms)` |
| AA `Avg Output Tokens` | `Avg Output Tokens` |
| AB `Input Throughput (tok/s)` | `Input Throughput (tok/s)` |

## 常用状态文件

完整运行输出在 `RunDir` 下：

```text
<RunDir>/
  logs/
  outputs/
    short/
    long/
  run-config.redacted.json
  status.json
  run-summary.md
  模型测试汇总表-new.xlsx
```

查看后台状态：

```powershell
Get-Content ".\llm-benchmark-runs\20260525-Qwen3.5-122B-A10B\status.json"
Get-Content ".\llm-benchmark-runs\20260525-Qwen3.5-122B-A10B\logs\run.log" -Tail 20
```

## 故障处理

- `benchmark_percentile.json` 缺失：该并发没有完整分位数输出，对应行不会填。
- `benchmark_summary.json` 缺失：该并发汇总列不会填。
- Excel 行没填上：检查 `excel_model_name` 是否和表内模型名一致。
- `Cannot run 'conda run -n ... evalscope --help'`：本机没有可用的 EvalScope conda 环境，或环境名不对。让用户手动准备环境，或用 `-EvalScopeCondaEnv` 指向已有环境。
- `status.json` 显示 `no output progress`：当前阶段输出目录长期没有变化，watchdog 已停止子进程。
- `run.log` 显示 `exited with code 1 but produced successful requests`：有部分请求失败，但已有可汇总成功数据，runner 会继续。
