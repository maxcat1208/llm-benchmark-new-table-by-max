# LLM Benchmark New Table by Max

Codex skill for running EvalScope LLM performance benchmarks and filling the two-sheet `模型测试汇总表-new.xlsx` workbook.

See [SKILL.md](SKILL.md) for the workflow, environment checks, runner usage, and Excel field mapping.

## Contents

- `SKILL.md`: Codex skill instructions.
- `scripts/run_full_benchmark.ps1`: Runs short then long EvalScope benchmarks and fills the workbook.
- `scripts/fill_new_benchmark_excel.py`: Fills an existing workbook from EvalScope outputs.
- `scripts/assert_evalscope_success.py`: Checks whether EvalScope produced successful requests.
- `templates/模型测试汇总表-new.xlsx`: Workbook template.
- `tests/`: Smoke tests for workbook filling.

## Safety

The runner checks for an existing EvalScope conda environment and does not install dependencies automatically. Logs and generated argument files are redacted for `sk-*` style API keys.
