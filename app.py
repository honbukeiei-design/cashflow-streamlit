import streamlit as st
from pathlib import Path
import tempfile

from cashflow_logic import run_cashflow_process

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

            with tempfile.TemporaryDirectory() as tmpdir:
                base_dir = Path(tmpdir)

                def save_uploaded_file(uploaded_file):
                    path = base_dir / uploaded_file.name
                    with open(path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    return path

                for file in voucher_files:
                    save_uploaded_file(file)

                save_uploaded_file(detail_master)
                save_uploaded_file(shozoku_master)
                save_uploaded_file(interface_master)
                save_uploaded_file(cashflow_items)

                if manual_file is not None:
                    save_uploaded_file(manual_file)

                try:
                    result = run_cashflow_process(base_dir)

                    st.success("処理が完了しました。")

                    st.write("### 処理結果")
                    st.write(f"伝票検索ファイル数：{result['voucher_file_count']:,}")
                    st.write(f"伝票検索取込件数：{result['voucher_row_count']:,}")
                    st.write(f"CF対象件数：{result['cf_row_count']:,}")
                    st.write(f"CF反映済件数：{result['matched_count']:,}")
                    st.write(f"未反映件数：{result['unmatched_count']:,}")
                    st.write(f"手入力反映件数：{result['manual_reflect_count']:,}")

                    with open(result["output_path"], "rb") as f:
                        st.download_button(
                            "cashflow_output.xlsx をダウンロード",
                            data=f,
                            file_name="cashflow_output.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                    with open(result["unmatched_path"], "rb") as f:
                        st.download_button(
                            "未反映仕訳一覧.xlsx をダウンロード",
                            data=f,
                            file_name="未反映仕訳一覧.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                    with open(result["manual_log_path"], "rb") as f:
                        st.download_button(
                            "手入力反映ログ.xlsx をダウンロード",
                            data=f,
                            file_name="手入力反映ログ.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                except Exception as e:
                    st.error("処理中にエラーが発生しました。")
                    st.exception(e)
