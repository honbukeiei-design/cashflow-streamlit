import streamlit as st
from pathlib import Path
import tempfile
import subprocess
import sys
import shutil

st.title("キャッシュフロー集計ツール")

st.write("必要なExcelファイルをアップロードして、集計を実行してください。")

voucher_files = st.file_uploader(
    "伝票検索ファイル（複数可）",
    type=["xlsx"],
    accept_multiple_files=True
)

detail_master = st.file_uploader("Q_010詳細科目一覧.xlsx", type=["xlsx"])
shozoku_master = st.file_uploader("T_所属コード.xlsx", type=["xlsx"])
interface_master = st.file_uploader("T_InterfaceMaster.xlsx", type=["xlsx"])
cashflow_items = st.file_uploader("T_Cashflow内訳項目.xlsx", type=["xlsx"])

manual_file = st.file_uploader(
    "未反映仕訳_手入力反映用.xlsx（任意）",
    type=["xlsx"]
)

correction_file = st.file_uploader(
    "CF反映済_修正用.xlsx（任意）",
    type=["xlsx"]
)

def save_uploaded_file(uploaded_file, folder: Path):
    path = folder / uploaded_file.name
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return path

if st.button("集計を実行"):

    if (
        not voucher_files
        or detail_master is None
        or shozoku_master is None
        or interface_master is None
        or cashflow_items is None
    ):
        st.error("必要ファイルが不足しています。")
    else:
        with st.spinner("処理中です..."):

            try:
                with tempfile.TemporaryDirectory() as tmpdir:
                    base_dir = Path(tmpdir)

                    for file in voucher_files:
                        save_uploaded_file(file, base_dir)

                    save_uploaded_file(detail_master, base_dir)
                    save_uploaded_file(shozoku_master, base_dir)
                    save_uploaded_file(interface_master, base_dir)
                    save_uploaded_file(cashflow_items, base_dir)

                    if manual_file is not None:
                        save_uploaded_file(manual_file, base_dir)

                    if correction_file is not None:
                        save_uploaded_file(correction_file, base_dir)

                    script_source = Path(__file__).resolve().parent / "cashflow_script.py"
                    script_target = base_dir / "cashflow_script.py"
                    shutil.copy(script_source, script_target)

                    completed = subprocess.run(
                        [sys.executable, str(script_target)],
                        cwd=base_dir,
                        capture_output=True,
                        text=True,
                        check=False
                    )

                    if completed.returncode != 0:
                        st.error("処理中にエラーが発生しました。")
                        st.code(completed.stderr)
                    else:
                        output_path = base_dir / "cashflow_output.xlsx"
                        unmatched_path = base_dir / "未反映仕訳一覧.xlsx"
                        manual_log_path = base_dir / "手入力反映ログ.xlsx"
                        correction_log_path = base_dir / "CF反映修正ログ.xlsx"

                        st.session_state["stdout"] = completed.stdout
                        st.session_state["output_file"] = output_path.read_bytes()
                        st.session_state["unmatched_file"] = unmatched_path.read_bytes()
                        st.session_state["manual_log_file"] = manual_log_path.read_bytes()
                        st.session_state["correction_log_file"] = correction_log_path.read_bytes()

                        st.success("処理が完了しました。下のボタンからダウンロードしてください。")

            except Exception as e:
                st.error("処理中にエラーが発生しました。")
                st.exception(e)

if "output_file" in st.session_state:

    st.write("### 実行結果")

    if "stdout" in st.session_state:
        st.code(st.session_state["stdout"])

    st.download_button(
        "cashflow_output.xlsx をダウンロード",
        data=st.session_state["output_file"],
        file_name="cashflow_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.download_button(
        "未反映仕訳一覧.xlsx をダウンロード",
        data=st.session_state["unmatched_file"],
        file_name="未反映仕訳一覧.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.download_button(
        "手入力反映ログ.xlsx をダウンロード",
        data=st.session_state["manual_log_file"],
        file_name="手入力反映ログ.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.download_button(
        "CF反映修正ログ.xlsx をダウンロード",
        data=st.session_state["correction_log_file"],
        file_name="CF反映修正ログ.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
