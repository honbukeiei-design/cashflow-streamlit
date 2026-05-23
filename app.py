import streamlit as st
from pathlib import Path
import tempfile

st.title("キャッシュフロー集計ツール")

st.write("必要なExcelファイルをアップロードしてください。")

# 伝票検索ファイル（複数可）
voucher_files = st.file_uploader(
    "伝票検索ファイル",
    type=["xlsx"],
    accept_multiple_files=True
)

# マスタファイル
detail_master = st.file_uploader(
    "Q_010詳細科目一覧.xlsx",
    type=["xlsx"]
)

shozoku_master = st.file_uploader(
    "T_所属コード.xlsx",
    type=["xlsx"]
)

interface_master = st.file_uploader(
    "T_InterfaceMaster.xlsx",
    type=["xlsx"]
)

cashflow_items = st.file_uploader(
    "T_Cashflow内訳項目.xlsx",
    type=["xlsx"]
)

manual_file = st.file_uploader(
    "未反映仕訳_手入力反映用.xlsx（任意）",
    type=["xlsx"]
)

# 実行ボタン
if st.button("実行"):

    # 必須チェック
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

            # 一時フォルダ作成
            with tempfile.TemporaryDirectory() as tmpdir:

                base_dir = Path(tmpdir)

                # ファイル保存関数
                def save_uploaded_file(uploaded_file):
                    path = base_dir / uploaded_file.name
                    with open(path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    return path

                # 伝票検索ファイル保存
                for file in voucher_files:
                    save_uploaded_file(file)

                # マスタ保存
                save_uploaded_file(detail_master)
                save_uploaded_file(shozoku_master)
                save_uploaded_file(interface_master)
                save_uploaded_file(cashflow_items)

                # 任意ファイル
                if manual_file is not None:
                    save_uploaded_file(manual_file)

                st.success("ファイルアップロード成功")

                st.write("ここに集計処理を追加します。")
