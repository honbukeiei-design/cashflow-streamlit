from pathlib import Path

def run_cashflow_process(base_dir: Path):
    # 今の処理本体をここに入れる
    return {
        "output_file": base_dir / "cashflow_output.xlsx",
        "unmatched_file": base_dir / "未反映仕訳一覧.xlsx",
        "manual_log_file": base_dir / "手入力反映ログ.xlsx",
    }
