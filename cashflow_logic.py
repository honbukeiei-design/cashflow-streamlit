from pathlib import Path
import pandas as pd
import numpy as np

# =========================================================
# 設定
# =========================================================

BASE_DIR = Path(__file__).resolve().parent

VOUCHER_KEYWORD = "伝票検索"

DETAIL_MASTER_FILE = "Q_010詳細科目一覧.xlsx"
SHOZOKU_MASTER_FILE = "T_所属コード.xlsx"
INTERFACE_MASTER_FILE = "T_InterfaceMaster.xlsx"
CASHFLOW_ITEM_FILE = "T_Cashflow内訳項目.xlsx"

OUTPUT_FILE = "cashflow_output.xlsx"
UNMATCHED_FILE = "未反映仕訳一覧.xlsx"

# 未反映仕訳を手入力でCFに反映するためのファイル
MANUAL_INPUT_FILE = "未反映仕訳_手入力反映用.xlsx"
MANUAL_LOG_FILE = "手入力反映ログ.xlsx"
MANUAL_CODE_COLUMN = "手入力項目コード"

# すでにCF反映済みの仕訳を修正するためのファイル
CORRECTION_INPUT_FILE = "CF反映済_修正用.xlsx"
CORRECTION_LOG_FILE = "CF反映修正ログ.xlsx"
CORRECTION_CODE_COLUMN = "修正項目コード"

CASH_ACCOUNT_PREFIX = "040202"


# =========================================================
# 共通関数
# =========================================================

def find_excel_files(keyword: str, folder: Path) -> list[Path]:
    files = [
        p for p in folder.glob("*.xlsx")
        if keyword in p.name and not p.name.startswith("~$")
    ]
    if not files:
        raise FileNotFoundError(f"'{keyword}' を含むExcelファイルが見つかりません。")
    return sorted(files)


def read_excel_as_text(path: Path) -> pd.DataFrame:
    return pd.read_excel(path, dtype=str).fillna("")


def clean_code(series: pd.Series) -> pd.Series:
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.replace(".0", "", regex=False)
        .replace(["nan", "NaN", "None", "<NA>"], "")
    )


def is_blank(series: pd.Series) -> pd.Series:
    return clean_code(series).isin(["", "nan", "NaN", "None", "<NA>"])


def to_number(series: pd.Series) -> pd.Series:
    return (
        series.fillna("")
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.strip()
        .replace("", "0")
        .astype(float)
    )


def require_columns(df: pd.DataFrame, columns: list[str], name: str) -> None:
    missing = [c for c in columns if c not in df.columns]
    if missing:
        raise ValueError(f"{name} に必要な列がありません: {missing}")


def apply_item_name(df: pd.DataFrame, cashflow_items: pd.DataFrame) -> pd.DataFrame:
    """
    項目コードから項目名を付与する。
    mergeを繰り返すと行数が増えるリスクがあるため、mapで付与する。
    """
    item_master = cashflow_items[["項目コード", "項目名"]].copy()
    item_master["項目コード"] = clean_code(item_master["項目コード"])
    item_master["項目名"] = clean_code(item_master["項目名"])

    item_master = item_master.drop_duplicates(subset=["項目コード"], keep="last")

    item_name_map = item_master.set_index("項目コード")["項目名"]

    df["項目コード"] = clean_code(df["項目コード"])
    df["項目名"] = df["項目コード"].map(item_name_map).fillna("")
    df["項目名"] = clean_code(df["項目名"])

    return df


# =========================================================
# データ読込
# =========================================================

voucher_files = find_excel_files(VOUCHER_KEYWORD, BASE_DIR)

print("取込対象ファイル:")
for f in voucher_files:
    print(f"  - {f.name}")

voucher_dfs = []

for file in voucher_files:
    temp_df = read_excel_as_text(file)
    temp_df["取込元ファイル"] = file.name
    voucher_dfs.append(temp_df)

t_a = pd.concat(voucher_dfs, ignore_index=True)

q_010 = read_excel_as_text(BASE_DIR / DETAIL_MASTER_FILE)
t_shozoku = read_excel_as_text(BASE_DIR / SHOZOKU_MASTER_FILE)
interface_master = read_excel_as_text(BASE_DIR / INTERFACE_MASTER_FILE)
cashflow_items = read_excel_as_text(BASE_DIR / CASHFLOW_ITEM_FILE)

print(f"伝票検索 取込件数: {len(t_a):,}")


# =========================================================
# 必須列チェック
# =========================================================

require_columns(
    t_a,
    [
        "所属",
        "伝票区分名称",
        "様式番号",
        "伝票番号",
        "行番号",
        "貸借区分",
        "勘定科目コード",
        "勘定科目名称",
        "発生科目コード",
        "発生科目名称",
        "相手科目コード",
        "相手科目名称",
        "税込額",
        "件名",
        "摘要",
        "仕訳日",
    ],
    "伝票検索データ",
)

require_columns(q_010, ["款項目節細", "款項目節", "節略称"], "Q_010詳細科目一覧")
require_columns(t_shozoku, ["所属名", "表示順", "所属コード名"], "T_所属コード")
require_columns(interface_master, ["ＣＦ仕訳フラグ", "項目コード"], "T_InterfaceMaster")
require_columns(cashflow_items, ["項目コード", "項目名"], "T_Cashflow内訳項目")


# =========================================================
# コード列の整形
# =========================================================

for col in [
    "所属",
    "様式番号",
    "伝票番号",
    "行番号",
    "貸借区分",
    "勘定科目コード",
    "発生科目コード",
    "相手科目コード",
]:
    t_a[col] = clean_code(t_a[col])

for col in ["款項目節細", "款項目節"]:
    q_010[col] = clean_code(q_010[col])

for col in ["ＣＦ仕訳フラグ", "項目コード"]:
    interface_master[col] = clean_code(interface_master[col])

cashflow_items["項目コード"] = clean_code(cashflow_items["項目コード"])
cashflow_items["項目名"] = clean_code(cashflow_items["項目名"])


# =========================================================
# sub1 相当：元データ一部加工
# =========================================================

df = t_a.copy()

df = df.merge(
    q_010[["款項目節細", "款項目節", "節略称"]],
    how="left",
    left_on="相手科目コード",
    right_on="款項目節細",
)

df = df.merge(
    t_shozoku[["所属名", "表示順", "所属コード名"]],
    how="left",
    left_on="所属",
    right_on="所属名",
)

df["相手科目節コード"] = clean_code(df["款項目節"])
df["相手科目節略称"] = clean_code(df["節略称"])

df["識別ID"] = (
    clean_code(df["表示順"])
    + clean_code(df["様式番号"])
    + clean_code(df["伝票番号"])
    + clean_code(df["行番号"])
    + clean_code(df["貸借区分"])
)

df["税込額_num"] = to_number(df["税込額"])

df["仕訳額"] = np.where(
    clean_code(df["貸借区分"]) == "L",
    df["税込額_num"],
    -df["税込額_num"],
)

df["相手科目コード先頭2桁"] = pd.to_numeric(
    clean_code(df["相手科目コード"]).str[:2],
    errors="coerce",
)

df["項目抽出対象"] = np.where(
    df["相手科目コード先頭2桁"] < 10,
    "発生",
    "相手",
)

df["勘定相手12桁"] = (
    clean_code(df["勘定科目コード"]).str[:6]
    + clean_code(df["相手科目コード"]).str[:6]
)


# =========================================================
# sub2 相当：対象データ抽出
# =========================================================

condition = (
    (clean_code(df["伝票区分名称"]) != "付替")
    & clean_code(df["勘定科目コード"]).str.startswith(CASH_ACCOUNT_PREFIX)
    & ~clean_code(df["相手科目コード"]).str.startswith(CASH_ACCOUNT_PREFIX)
)

df = df[condition].copy()
df["抽出項目"] = df["項目抽出対象"]


# =========================================================
# sub3 相当：CF仕訳科目抽出
# =========================================================

df["ＣＦ仕訳科目コード"] = np.where(
    df["抽出項目"] == "相手",
    df["相手科目節コード"],
    df["発生科目コード"],
)

df["ＣＦ仕訳科目名"] = np.where(
    df["抽出項目"] == "相手",
    df["相手科目節略称"],
    df["発生科目名称"],
)

df["ＣＦ仕訳科目コード"] = clean_code(df["ＣＦ仕訳科目コード"])

df["ＣＦ仕訳フラグ"] = (
    clean_code(df["ＣＦ仕訳科目コード"])
    + clean_code(df["貸借区分"])
)


# =========================================================
# 通常マスタ結合
# =========================================================

df = df.merge(
    interface_master[["ＣＦ仕訳フラグ", "項目コード"]],
    how="left",
    on="ＣＦ仕訳フラグ",
)

df["項目コード"] = clean_code(df["項目コード"])
df = apply_item_name(df, cashflow_items)


# =========================================================
# 未反映理由の初期判定
# =========================================================

df["未反映理由"] = ""

blank_accrual_code = (
    (df["抽出項目"] == "発生")
    & is_blank(df["発生科目コード"])
)

blank_partner_section = (
    (df["抽出項目"] == "相手")
    & is_blank(df["相手科目節コード"])
)

blank_cf_code = is_blank(df["ＣＦ仕訳科目コード"])
blank_item_code = is_blank(df["項目コード"])
blank_item_name = is_blank(df["項目名"])

df.loc[blank_cf_code, "未反映理由"] = "CF仕訳科目コードが空白"
df.loc[blank_accrual_code, "未反映理由"] = "発生科目コードが空白"
df.loc[blank_partner_section, "未反映理由"] = "相手科目節コードが空白"

df.loc[
    blank_item_code & is_blank(df["未反映理由"]),
    "未反映理由",
] = "T_InterfaceMasterに未登録"

df.loc[
    (~blank_item_code) & blank_item_name & is_blank(df["未反映理由"]),
    "未反映理由",
] = "T_Cashflow内訳項目に未登録"


# =========================================================
# 未反映仕訳の手入力反映
# 取込ファイル：未反映仕訳_手入力反映用.xlsx
# =========================================================

manual_reflect_log = pd.DataFrame()
manual_path = BASE_DIR / MANUAL_INPUT_FILE

if manual_path.exists():
    manual_df = read_excel_as_text(manual_path)

    if MANUAL_CODE_COLUMN in manual_df.columns:
        require_columns(
            manual_df,
            ["取込元ファイル", "識別ID", MANUAL_CODE_COLUMN],
            MANUAL_INPUT_FILE,
        )

        manual_df["取込元ファイル"] = clean_code(manual_df["取込元ファイル"])
        manual_df["識別ID"] = clean_code(manual_df["識別ID"])
        manual_df[MANUAL_CODE_COLUMN] = clean_code(manual_df[MANUAL_CODE_COLUMN])

        manual_df = manual_df[
            ~is_blank(manual_df[MANUAL_CODE_COLUMN])
        ][["取込元ファイル", "識別ID", MANUAL_CODE_COLUMN]]

        manual_df = manual_df.drop_duplicates(
            subset=["取込元ファイル", "識別ID"],
            keep="last",
        )

        df = df.merge(
            manual_df,
            how="left",
            on=["取込元ファイル", "識別ID"],
        )

        manual_reflect_condition = (
            ~is_blank(df[MANUAL_CODE_COLUMN])
            & (
                is_blank(df["項目コード"])
                | is_blank(df["項目名"])
                | df["未反映理由"].isin(
                    [
                        "相手科目節コードが空白",
                        "発生科目コードが空白",
                        "CF仕訳科目コードが空白",
                        "T_InterfaceMasterに未登録",
                        "T_Cashflow内訳項目に未登録",
                    ]
                )
            )
        )

        df.loc[manual_reflect_condition, "項目コード"] = clean_code(
            df.loc[manual_reflect_condition, MANUAL_CODE_COLUMN]
        )

        df = apply_item_name(df, cashflow_items)

        df.loc[manual_reflect_condition, "未反映理由"] = ""

        invalid_manual_code = (
            manual_reflect_condition
            & is_blank(df["項目名"])
        )

        df.loc[invalid_manual_code, "未反映理由"] = "手入力項目コードが内訳項目マスタに未登録"

        manual_reflect_log = df.loc[
            manual_reflect_condition & ~invalid_manual_code
        ].copy()

        print(f"{MANUAL_INPUT_FILE} を読み込みました。")
    else:
        df[MANUAL_CODE_COLUMN] = ""
        print(f"{MANUAL_INPUT_FILE} に {MANUAL_CODE_COLUMN} 列がないため、手入力反映はスキップしました。")
else:
    df[MANUAL_CODE_COLUMN] = ""
    print(f"{MANUAL_INPUT_FILE} がないため、手入力反映なしで処理しました。")


# =========================================================
# 反映済仕訳の修正反映
# 取込ファイル：CF反映済_修正用.xlsx
#
# 今回の構成では、以下の3列だけ必須
#   取込元ファイル
#   識別ID
#   修正項目コード
# =========================================================

correction_log = pd.DataFrame()
correction_path = BASE_DIR / CORRECTION_INPUT_FILE

if correction_path.exists():
    correction_df = read_excel_as_text(correction_path)

    if CORRECTION_CODE_COLUMN in correction_df.columns:
        require_columns(
            correction_df,
            ["取込元ファイル", "識別ID", CORRECTION_CODE_COLUMN],
            CORRECTION_INPUT_FILE,
        )

        correction_df["取込元ファイル"] = clean_code(correction_df["取込元ファイル"])
        correction_df["識別ID"] = clean_code(correction_df["識別ID"])
        correction_df[CORRECTION_CODE_COLUMN] = clean_code(
            correction_df[CORRECTION_CODE_COLUMN]
        )

        correction_df = correction_df[
            ~is_blank(correction_df[CORRECTION_CODE_COLUMN])
        ][["取込元ファイル", "識別ID", CORRECTION_CODE_COLUMN]]

        correction_df = correction_df.drop_duplicates(
            subset=["取込元ファイル", "識別ID"],
            keep="last",
        )

        df = df.merge(
            correction_df,
            how="left",
            on=["取込元ファイル", "識別ID"],
        )

        correction_condition = ~is_blank(df[CORRECTION_CODE_COLUMN])

        df["修正前項目コード"] = df["項目コード"]
        df["修正前項目名"] = df["項目名"]
        df["修正前未反映理由"] = df["未反映理由"]

        df.loc[correction_condition, "項目コード"] = clean_code(
            df.loc[correction_condition, CORRECTION_CODE_COLUMN]
        )

        df = apply_item_name(df, cashflow_items)

        df.loc[correction_condition, "未反映理由"] = ""

        invalid_correction_code = (
            correction_condition
            & is_blank(df["項目名"])
        )

        df.loc[
            invalid_correction_code,
            "未反映理由",
        ] = "修正項目コードが内訳項目マスタに未登録"

        correction_log = df.loc[correction_condition].copy()

        correction_log["修正結果"] = np.where(
            is_blank(correction_log["項目名"]),
            "修正項目コードが内訳項目マスタに未登録",
            "修正反映済",
        )

        print(f"{CORRECTION_INPUT_FILE} を読み込みました。")
    else:
        df[CORRECTION_CODE_COLUMN] = ""
        print(f"{CORRECTION_INPUT_FILE} に {CORRECTION_CODE_COLUMN} 列がないため、修正反映はスキップしました。")
else:
    df[CORRECTION_CODE_COLUMN] = ""
    print(f"{CORRECTION_INPUT_FILE} がないため、修正反映なしで処理しました。")


# =========================================================
# 反映済・未反映に分離
# =========================================================

df["項目コード"] = clean_code(df["項目コード"])
df["項目名"] = clean_code(df["項目名"])

unmatched_condition = (
    is_blank(df["項目コード"])
    | is_blank(df["項目名"])
)

df.loc[
    unmatched_condition & is_blank(df["未反映理由"]),
    "未反映理由",
] = "項目コードまたは項目名が空白"

unmatched = df[unmatched_condition].copy()
matched = df[~unmatched_condition].copy()


# =========================================================
# CF集計結果
# =========================================================

result = (
    matched
    .groupby(["項目コード", "項目名"], dropna=False)["仕訳額"]
    .sum()
    .reset_index(name="仕訳額の合計")
    .sort_values("項目コード")
)


# =========================================================
# 未登録マスタ候補
# =========================================================

missing_master_summary = (
    unmatched
    .groupby(
        [
            "未反映理由",
            "ＣＦ仕訳フラグ",
            "ＣＦ仕訳科目コード",
            "ＣＦ仕訳科目名",
            "貸借区分",
        ],
        dropna=False,
    )["仕訳額"]
    .agg(["count", "sum"])
    .reset_index()
    .rename(columns={"count": "件数", "sum": "仕訳額合計"})
)


# =========================================================
# 未反映仕訳一覧の出力列
# =========================================================

if MANUAL_CODE_COLUMN not in unmatched.columns:
    unmatched[MANUAL_CODE_COLUMN] = ""

if CORRECTION_CODE_COLUMN not in unmatched.columns:
    unmatched[CORRECTION_CODE_COLUMN] = ""

unmatched_columns = [
    "未反映理由",
    "取込元ファイル",
    "識別ID",
    "仕訳日",
    "表示順",
    "所属",
    "所属コード名",
    "伝票区分名称",
    "様式番号",
    "伝票番号",
    "行番号",
    "貸借区分",
    "勘定科目コード",
    "勘定科目名称",
    "発生科目コード",
    "発生科目名称",
    "相手科目コード",
    "相手科目名称",
    "相手科目節コード",
    "相手科目節略称",
    "抽出項目",
    "ＣＦ仕訳科目コード",
    "ＣＦ仕訳科目名",
    "ＣＦ仕訳フラグ",
    "仕訳額",
    "件名",
    "摘要",
    MANUAL_CODE_COLUMN,
]

unmatched_columns = [c for c in unmatched_columns if c in unmatched.columns]


# =========================================================
# 手入力反映ログの出力列
# =========================================================

manual_log_columns = [
    "取込元ファイル",
    "識別ID",
    "仕訳日",
    "表示順",
    "所属",
    "所属コード名",
    "伝票区分名称",
    "様式番号",
    "伝票番号",
    "行番号",
    "貸借区分",
    "勘定科目コード",
    "勘定科目名称",
    "発生科目コード",
    "発生科目名称",
    "相手科目コード",
    "相手科目名称",
    "相手科目節コード",
    "相手科目節略称",
    "抽出項目",
    "ＣＦ仕訳科目コード",
    "ＣＦ仕訳科目名",
    "ＣＦ仕訳フラグ",
    "仕訳額",
    "件名",
    "摘要",
    MANUAL_CODE_COLUMN,
    "項目コード",
    "項目名",
]

manual_log_columns = [
    c for c in manual_log_columns
    if c in manual_reflect_log.columns
]


# =========================================================
# 修正ログの出力列
# CF反映済_修正用.xlsx の構成に寄せる
# =========================================================

correction_log_columns = [
    "修正結果",
    "件名",
    "摘要",
    "取込元ファイル",
    "款項目節細",
    "款項目節",
    "節略称",
    "所属名",
    "表示順",
    "所属コード名",
    "相手科目節コード",
    "相手科目節略称",
    "識別ID",
    "税込額_num",
    "仕訳額",
    "相手科目コード先頭2桁",
    "項目抽出対象",
    "勘定相手12桁",
    "抽出項目",
    "ＣＦ仕訳科目コード",
    "ＣＦ仕訳科目名",
    "ＣＦ仕訳フラグ",
    "修正前項目コード",
    "修正前項目名",
    CORRECTION_CODE_COLUMN,
    "項目コード",
    "項目名",
    "修正前未反映理由",
    "未反映理由",
]

correction_log_columns = [
    c for c in correction_log_columns
    if c in correction_log.columns
]


# =========================================================
# Excel出力
# =========================================================

output_path = BASE_DIR / OUTPUT_FILE
unmatched_path = BASE_DIR / UNMATCHED_FILE
manual_log_path = BASE_DIR / MANUAL_LOG_FILE
correction_log_path = BASE_DIR / CORRECTION_LOG_FILE

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    result.to_excel(writer, sheet_name="CF集計結果", index=False)
    missing_master_summary.to_excel(writer, sheet_name="未登録マスタ候補", index=False)
    matched.to_excel(writer, sheet_name="反映済明細", index=False)

with pd.ExcelWriter(unmatched_path, engine="openpyxl") as writer:
    unmatched[unmatched_columns].to_excel(
        writer,
        sheet_name="未反映仕訳一覧",
        index=False,
    )

with pd.ExcelWriter(manual_log_path, engine="openpyxl") as writer:
    manual_reflect_log[manual_log_columns].to_excel(
        writer,
        sheet_name="手入力反映ログ",
        index=False,
    )

with pd.ExcelWriter(correction_log_path, engine="openpyxl") as writer:
    correction_log[correction_log_columns].to_excel(
        writer,
        sheet_name="CF反映修正ログ",
        index=False,
    )


# =========================================================
# 実行結果表示
# =========================================================

print("処理が完了しました。")
print(f"出力ファイル: {output_path}")
print(f"未反映仕訳一覧: {unmatched_path}")
print(f"手入力反映用ファイル: {BASE_DIR / MANUAL_INPUT_FILE}")
print(f"手入力反映ログ: {manual_log_path}")
print(f"CF反映済修正用ファイル: {BASE_DIR / CORRECTION_INPUT_FILE}")
print(f"CF反映修正ログ: {correction_log_path}")
print(f"伝票検索ファイル数: {len(voucher_files):,}")
print(f"伝票検索取込件数: {len(t_a):,}")
print(f"CF対象件数: {len(df):,}")
print(f"CF反映済件数: {len(matched):,}")
print(f"未反映件数: {len(unmatched):,}")
print(f"手入力反映件数: {len(manual_reflect_log):,}")
print(f"CF反映修正件数: {len(correction_log):,}")
