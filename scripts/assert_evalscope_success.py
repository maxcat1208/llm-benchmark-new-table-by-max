from __future__ import annotations

import argparse
import json
import sqlite3
import sys
from pathlib import Path


def inspect_success(source_dir: Path) -> dict:
    total = 0
    success = 0
    prompt_tokens = set()
    by_dir = {}
    for db_path in sorted(source_dir.glob("parallel_*_number_*/benchmark_data.db")):
        with sqlite3.connect(db_path) as con:
            row_total, row_success = con.execute(
                "select count(*), coalesce(sum(success), 0) from result"
            ).fetchone()
            tokens = [
                row[0]
                for row in con.execute(
                    "select distinct prompt_tokens from result where success = 1 and prompt_tokens is not null"
                ).fetchall()
            ]
        total += int(row_total)
        success += int(row_success)
        prompt_tokens.update(int(token) for token in tokens if token is not None)
        by_dir[db_path.parent.name] = {
            "total": int(row_total),
            "success": int(row_success),
            "prompt_tokens": sorted(int(token) for token in tokens if token is not None),
        }
    return {
        "source_dir": str(source_dir),
        "total": total,
        "success": success,
        "prompt_tokens": sorted(prompt_tokens),
        "by_dir": by_dir,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-dir", required=True)
    args = parser.parse_args()
    result = inspect_success(Path(args.source_dir))
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if result["success"] <= 0:
        sys.exit(2)


if __name__ == "__main__":
    main()
