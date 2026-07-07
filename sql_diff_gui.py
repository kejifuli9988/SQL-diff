import os
import re
import sys
import traceback
import hashlib
from zipfile import BadZipFile
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set, Tuple

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


SQL_COLUMN_NAME = "SQL语句"
TABLE_STATS_NAME_COLUMN = "表名"
TABLE_STATS_ROWS_COLUMN = "记录数"
HISTORY_ENRICH_COLUMNS = [
    "服务",
    "大表且暂不优化",
    "大表表名",
    "慢SQL分类",
    "初步优化方案",
    "应用场景",
    "加权分数",
    "优先级",
    "修复时间",
    "跟进情况",
    "备注",
    "表拆分后的平均执行时间",
]


def normalize_sql(value: object, ignore_whitespace: bool = True) -> str:
    if pd.isna(value):
        return ""

    text = str(value).replace("\u3000", " ").strip()
    if ignore_whitespace:
        text = re.sub(r"\s+", " ", text)
    return text


def classify_sql(sql: str) -> str:
    s = normalize_sql(sql).lower()
    if s.startswith("select count(*) from"):
        return "COUNT查询"
    if s.startswith("select * from ( select row_.*, rownum as rownum_ from"):
        return "分页查询"
    if s.startswith("select") and " for update" in s:
        return "SELECT FOR UPDATE"
    if s.startswith("select"):
        return "普通SELECT"
    if s.startswith("insert into"):
        return "INSERT"
    if s.startswith("update"):
        return "UPDATE"
    if s.startswith("delete from"):
        return "DELETE"
    if s.startswith("with"):
        return "WITH查询"
    if s.startswith("begin"):
        return "存储过程/PLSQL"
    return "其他"


def build_similarity_signature(sql: str) -> str:
    s = normalize_sql(sql).lower()
    s = re.sub(r"'(?:''|[^'])*'", "?str?", s)
    s = re.sub(r"\b\d{4}[-/]\d{1,2}[-/]\d{1,2}(?: \d{1,2}:\d{1,2}:\d{1,2})?\b", "?date?", s)
    s = re.sub(r"\b\d{8}\b", "?date8?", s)
    s = re.sub(r"\b\d+\b", "?num?", s)
    s = re.sub(r"\bin\s*\((?:[^()]*?)\)", "in(?list?)", s)
    s = re.sub(r"\bvalues\s*\((?:[^()]*?)\)", "values(?vals?)", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def build_date_only_signature(sql: str) -> str:
    s = normalize_sql(sql).lower()
    s = re.sub(r"\b\d{4}[-/]\d{1,2}[-/]\d{1,2}(?: \d{1,2}:\d{1,2}:\d{1,2})?\b", "?date?", s)
    s = re.sub(r"\b\d{8}\b", "?date8?", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def build_inlist_signature(sql: str) -> str:
    s = normalize_sql(sql).lower()
    s = re.sub(r"\bin\s*\((?:[^()]*?)\)", "in(?list?)", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def build_parameter_signature(sql: str) -> str:
    s = normalize_sql(sql).lower()
    s = re.sub(r"'(?:''|[^'])*'", "?str?", s)
    s = re.sub(r"\b\d{4}[-/]\d{1,2}[-/]\d{1,2}(?: \d{1,2}:\d{1,2}:\d{1,2})?\b", "?date?", s)
    s = re.sub(r"\b\d{8}\b", "?date8?", s)
    s = re.sub(r"\b\d+\b", "?num?", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def load_large_table_map(file_path: str, threshold: int) -> Dict[str, int]:
    df = pd.read_excel(file_path)
    if TABLE_STATS_NAME_COLUMN not in df.columns or TABLE_STATS_ROWS_COLUMN not in df.columns:
        raise ValueError(f"表数据量统计文件必须包含列：{TABLE_STATS_NAME_COLUMN}、{TABLE_STATS_ROWS_COLUMN}")

    result: Dict[str, int] = {}
    for _, row in df.iterrows():
        table_name = str(row.get(TABLE_STATS_NAME_COLUMN, "")).strip()
        row_count = row.get(TABLE_STATS_ROWS_COLUMN)
        if not table_name or pd.isna(row_count):
            continue
        try:
            count_int = int(float(row_count))
        except Exception:
            continue
        if count_int >= threshold:
            result[table_name.upper()] = count_int
    return result


def extract_table_names(sql: str) -> List[str]:
    normalized = normalize_sql(sql, ignore_whitespace=True)
    patterns = [
        r"\bfrom\s+([a-zA-Z0-9_.$]+)",
        r"\bjoin\s+([a-zA-Z0-9_.$]+)",
        r"\bupdate\s+([a-zA-Z0-9_.$]+)",
        r"\binsert\s+into\s+([a-zA-Z0-9_.$]+)",
        r"\bdelete\s+from\s+([a-zA-Z0-9_.$]+)",
        r"\bmerge\s+into\s+([a-zA-Z0-9_.$]+)",
    ]
    names: List[str] = []
    seen: Set[str] = set()
    for pattern in patterns:
        for match in re.findall(pattern, normalized, flags=re.IGNORECASE):
            name = str(match).strip().split(".")[-1].upper()
            if not name or name in {"SELECT", "DUAL"}:
                continue
            if name not in seen:
                seen.add(name)
                names.append(name)
    return names


def merge_group_values(series: pd.Series) -> str:
    values: List[str] = []
    seen: Set[str] = set()
    for value in series:
        if pd.isna(value):
            continue
        text = str(value).strip()
        if not text or text.lower() == "nan":
            continue
        if text not in seen:
            seen.add(text)
            values.append(text)
    return " | ".join(values)


def infer_cross_file_reason(grp: pd.DataFrame) -> Tuple[str, str]:
    if grp["来源文件"].nunique() < 2:
        return "否", "仅单表出现"

    file_names = list(grp["来源文件"].drop_duplicates())
    file_a = grp[grp["来源文件"] == file_names[0]]
    file_b = grp[grp["来源文件"] == file_names[1]]

    exact_a = set(file_a["SQL语句"])
    exact_b = set(file_b["SQL语句"])
    exact_overlap = exact_a & exact_b
    if exact_overlap:
        only_a = exact_a - exact_b
        only_b = exact_b - exact_a
        if not only_a and not only_b:
            return "是", "跨表完全相同"
        return "是", "跨表部分完全相同"

    date_a = set(file_a["日期归一特征"])
    date_b = set(file_b["日期归一特征"])
    if date_a & date_b:
        return "否", "仅日期不同"

    in_a = set(file_a["IN归一特征"])
    in_b = set(file_b["IN归一特征"])
    if in_a & in_b:
        return "否", "IN列表不同"

    param_a = set(file_a["参数归一特征"])
    param_b = set(file_b["参数归一特征"])
    if param_a & param_b:
        return "否", "仅参数不同"

    return "否", "结构相似"


def build_summary_row(
    grp: pd.DataFrame,
    class_id: str,
    cross_reason: str,
    has_exact_match: str,
    large_table_map: Optional[Dict[str, int]] = None,
) -> Dict[str, object]:
    positions = []
    for file_name, sub in grp.groupby("来源文件", sort=False):
        seqs = "、".join(str(x) for x in sub["对应表内第几条"].tolist())
        positions.append(f"{file_name}: {seqs}")

    matched_tables: List[str] = []
    if large_table_map:
        seen: Set[str] = set()
        for sql in grp["SQL语句"].astype(str):
            for table_name in extract_table_names(sql):
                if table_name in large_table_map and table_name not in seen:
                    seen.add(table_name)
                    matched_tables.append(table_name)

    return {
        "相似类ID": class_id,
        "SQL类型": grp["SQL类型"].iloc[0],
        "重复条数": len(grp),
        "涉及文件数": grp["来源文件"].nunique(),
        "是否存在完全相同SQL": has_exact_match,
        "跨表原因": cross_reason,
        "来源文件": "、".join(grp["来源文件"].drop_duplicates().tolist()),
        "对应表内第几条": " | ".join(positions),
        "代表SQL": grp["SQL语句"].iloc[0],
        "服务": merge_group_values(grp["服务"]) if "服务" in grp.columns else "",
        "大表且暂不优化": merge_group_values(grp["大表且暂不优化"]) if "大表且暂不优化" in grp.columns else "",
        "大表表名": merge_group_values(grp["大表表名"]) if "大表表名" in grp.columns else "",
        "慢SQL分类": merge_group_values(grp["慢SQL分类"]) if "慢SQL分类" in grp.columns else "",
        "初步优化方案": merge_group_values(grp["初步优化方案"]) if "初步优化方案" in grp.columns else "",
        "应用场景": merge_group_values(grp["应用场景"]) if "应用场景" in grp.columns else "",
        "加权分数": merge_group_values(grp["加权分数"]) if "加权分数" in grp.columns else "",
        "优先级": merge_group_values(grp["优先级"]) if "优先级" in grp.columns else "",
        "修复时间": merge_group_values(grp["修复时间"]) if "修复时间" in grp.columns else "",
        "跟进情况": merge_group_values(grp["跟进情况"]) if "跟进情况" in grp.columns else "",
        "备注": merge_group_values(grp["备注"]) if "备注" in grp.columns else "",
        "规则判断是否有大表": "是" if matched_tables else "否",
        "规则判断涉及到的大表名称": "、".join(matched_tables),
    }


def clean_col_name(value: object) -> str:
    return str(value).replace("\n", "").replace("\r", "").replace(" ", "").strip()


def load_excel(file_path: str) -> pd.DataFrame:
    suffix = Path(file_path).suffix.lower()
    if suffix == ".xlsx":
        return pd.read_excel(file_path, engine="openpyxl", header=None)
    if suffix == ".xlsm":
        return pd.read_excel(file_path, engine="openpyxl", header=None)
    if suffix == ".xls":
        try:
            return pd.read_excel(file_path, engine="xlrd", header=None)
        except Exception:
            tables = pd.read_html(file_path, header=None)
            return tables[0]

    raise ValueError(f"不支持的文件类型: {suffix}")


def fix_header(df: pd.DataFrame, target_col: str = SQL_COLUMN_NAME) -> Tuple[pd.DataFrame, int]:
    target = clean_col_name(target_col)

    for i in range(min(30, len(df))):
        row = [clean_col_name(x) for x in df.iloc[i].tolist()]
        if target in row:
            new_df = df.iloc[i + 1 :].copy()
            new_df.columns = row
            new_df.reset_index(drop=True, inplace=True)
            return new_df, i + 2

    raise ValueError(f"没有找到表头 “{target_col}”")


def build_sql_map(
    df: pd.DataFrame,
    ignore_whitespace: bool,
) -> Tuple[Dict[str, pd.Series], Iterable[str]]:
    sql_map: Dict[str, pd.Series] = {}
    ordered_keys = []

    for _, row in df.iterrows():
        normalized = normalize_sql(row.get(SQL_COLUMN_NAME), ignore_whitespace)
        if not normalized:
            continue
        if normalized not in sql_map:
            sql_map[normalized] = row
            ordered_keys.append(normalized)

    return sql_map, ordered_keys


def build_dataframe_from_rows(rows: Iterable[pd.Series], columns: Iterable[str]) -> pd.DataFrame:
    rows = list(rows)
    if not rows:
        return pd.DataFrame(columns=list(columns))
    return pd.DataFrame(rows, columns=list(columns))


def build_similarity_reports(
    file1: str,
    file2: str,
    df1: pd.DataFrame,
    df2: pd.DataFrame,
    start_row1: int,
    start_row2: int,
    large_table_map: Optional[Dict[str, int]] = None,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    detail_rows: List[Dict[str, object]] = []
    files = [(Path(file1).name, df1, start_row1), (Path(file2).name, df2, start_row2)]

    for file_name, df, start_row in files:
        for idx, row in df.reset_index(drop=True).iterrows():
            sql = normalize_sql(row.get(SQL_COLUMN_NAME))
            if not sql:
                continue
            strict_signature = build_similarity_signature(sql)
            date_signature = build_date_only_signature(sql)
            inlist_signature = build_inlist_signature(sql)
            parameter_signature = build_parameter_signature(sql)
            strict_class_id = "S" + hashlib.md5(strict_signature.encode("utf-8")).hexdigest()[:8].upper()
            raw_fingerprint = row.get("指纹", "") if "指纹" in df.columns else ""
            detail_rows.append(
                {
                    "来源文件": file_name,
                    "对应表内第几条": idx + 1,
                    "原始Excel行号": start_row + idx,
                    "相似类ID": strict_class_id,
                    "SQL类型": classify_sql(sql),
                    "原表指纹": raw_fingerprint,
                    "SQL语句": sql,
                    "相似SQL特征": strict_signature,
                    "日期归一特征": date_signature,
                    "IN归一特征": inlist_signature,
                    "参数归一特征": parameter_signature,
                    "SQL涉及表": "、".join(extract_table_names(sql)),
                }
            )
            for col in HISTORY_ENRICH_COLUMNS:
                detail_rows[-1][col] = row.get(col, "") if col in df.columns else ""

    detail_df = pd.DataFrame(detail_rows)
    if detail_df.empty:
        return pd.DataFrame(), pd.DataFrame()

    summary_rows = []
    for class_id, grp in detail_df.groupby("相似类ID", sort=False):
        if grp["来源文件"].nunique() < 2:
            summary_rows.append(build_summary_row(grp, class_id, "仅单表出现", "否", large_table_map))
            continue

        file_names = list(grp["来源文件"].drop_duplicates())
        file_a = grp[grp["来源文件"] == file_names[0]]
        file_b = grp[grp["来源文件"] == file_names[1]]
        exact_overlap = set(file_a["SQL语句"].astype(str)) & set(file_b["SQL语句"].astype(str))

        consumed_indexes: Set[int] = set()
        if exact_overlap:
            exact_grp = grp[grp["SQL语句"].astype(str).isin(exact_overlap)].copy()
            if not exact_grp.empty:
                summary_rows.append(build_summary_row(exact_grp, class_id, "跨表完全相同", "是", large_table_map))
                consumed_indexes.update(set(exact_grp.index.tolist()))

        remaining_grp = grp.loc[~grp.index.isin(consumed_indexes)].copy()
        if not remaining_grp.empty:
            if exact_overlap:
                summary_rows.append(build_summary_row(remaining_grp, class_id, "单表内完全相同", "否", large_table_map))
            else:
                has_exact_match, cross_reason = infer_cross_file_reason(remaining_grp)
                summary_rows.append(build_summary_row(remaining_grp, class_id, cross_reason, has_exact_match, large_table_map))

    summary_df = pd.DataFrame(summary_rows).sort_values(
        ["重复条数", "涉及文件数"],
        ascending=[False, False],
    )

    return summary_df, detail_df


def compare_sql_files(
    file1: str,
    file2: str,
    output_dir: str,
    ignore_whitespace: bool = True,
    deduplicate_within_file: bool = False,
    table_stats_file: str = "",
    large_table_threshold: int = 1000000,
) -> Tuple[str, Dict[str, int], pd.DataFrame, pd.DataFrame]:
    os.makedirs(output_dir, exist_ok=True)

    df1 = load_excel(file1)
    df2 = load_excel(file2)

    df1, start_row1 = fix_header(df1, SQL_COLUMN_NAME)
    df2, start_row2 = fix_header(df2, SQL_COLUMN_NAME)

    df1 = df1[df1[SQL_COLUMN_NAME].notna()].copy()
    df2 = df2[df2[SQL_COLUMN_NAME].notna()].copy()

    df1["_compare_sql_key"] = df1[SQL_COLUMN_NAME].apply(lambda x: normalize_sql(x, ignore_whitespace))
    df2["_compare_sql_key"] = df2[SQL_COLUMN_NAME].apply(lambda x: normalize_sql(x, ignore_whitespace))

    if deduplicate_within_file:
        df1 = df1.drop_duplicates(subset=["_compare_sql_key"], keep="first").copy()
        df2 = df2.drop_duplicates(subset=["_compare_sql_key"], keep="first").copy()
    else:
        # Keep the helper column for non-deduplicated filtering, then drop it from exported sheets later.
        pass

    keys1 = set(df1["_compare_sql_key"])
    keys2 = set(df2["_compare_sql_key"])
    only1_keys = keys1 - keys2
    only2_keys = keys2 - keys1
    both_keys = keys1 & keys2

    only1_df = df1[df1["_compare_sql_key"].isin(only1_keys)].drop(columns=["_compare_sql_key"])
    only2_df = df2[df2["_compare_sql_key"].isin(only2_keys)].drop(columns=["_compare_sql_key"])
    both1_df = df1[df1["_compare_sql_key"].isin(both_keys)].drop(columns=["_compare_sql_key"])
    both2_df = df2[df2["_compare_sql_key"].isin(both_keys)].drop(columns=["_compare_sql_key"])

    df1 = df1.drop(columns=["_compare_sql_key"])
    df2 = df2.drop(columns=["_compare_sql_key"])

    name1 = Path(file1).stem
    name2 = Path(file2).stem
    output_path = os.path.join(output_dir, f"SQL比较结果_{name1}_VS_{name2}.xlsx")

    sheet_only1 = f"仅{name1}"[:31]
    sheet_only2 = f"仅{name2}"[:31]
    sheet_both1 = f"共有({name1})"[:31]
    sheet_both2 = f"共有({name2})"[:31]
    large_table_map: Optional[Dict[str, int]] = None
    if table_stats_file.strip():
        large_table_map = load_large_table_map(table_stats_file.strip(), large_table_threshold)

    similarity_summary_df, similarity_detail_df = build_similarity_reports(
        file1=file1,
        file2=file2,
        df1=df1,
        df2=df2,
        start_row1=start_row1,
        start_row2=start_row2,
        large_table_map=large_table_map,
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df1.to_excel(writer, sheet_name=f"原表_{name1}"[:31], index=False)
        df2.to_excel(writer, sheet_name=f"原表_{name2}"[:31], index=False)
        only1_df.to_excel(writer, sheet_name=sheet_only1, index=False)
        only2_df.to_excel(writer, sheet_name=sheet_only2, index=False)
        both1_df.to_excel(writer, sheet_name=sheet_both1, index=False)
        both2_df.to_excel(writer, sheet_name=sheet_both2, index=False)
        similarity_summary_df.to_excel(writer, sheet_name="相似SQL归类汇总", index=False)
        similarity_detail_df.to_excel(writer, sheet_name="相似SQL明细", index=False)

    stats = {
        "表1总数": len(df1),
        "表2总数": len(df2),
        "仅表1": len(only1_df),
        "仅表2": len(only2_df),
        "共有": len(both1_df),
        "相似类数": len(similarity_summary_df),
        "大表阈值": large_table_threshold,
        "表内去重": "是" if deduplicate_within_file else "否",
    }
    return output_path, stats, similarity_summary_df, similarity_detail_df


class SqlDiffApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("SQL 语句 Excel 比较工具")
        self.root.geometry("680x630")
        self.root.minsize(600, 400)

        self.file1_var = tk.StringVar()
        self.file2_var = tk.StringVar()
        self.table_stats_var = tk.StringVar()
        self.output_dir_var = tk.StringVar(value=str(Path.home() / "Desktop"))
        self.large_table_threshold_var = tk.StringVar(value="1000000")
        self.relaxed_match_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="请选择两个 Excel 文件和输出目录。")

        self._build_ui()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=20)
        frame.pack(fill="both", expand=True)

        title = ttk.Label(frame, text="SQL 语句 Excel 比较工具", font=("Microsoft YaHei UI", 16, "bold"))
        title.pack(anchor="w")

        desc = ttk.Label(
            frame,
            text="比较两个 Excel 中“SQL语句”列，输出表1独有、表2独有，以及两个表都有的结果。",
            foreground="#444444",
        )
        desc.pack(anchor="w", pady=(8, 20))

        self._build_file_row(frame, "表1 Excel", self.file1_var, self.choose_file1)
        self._build_file_row(frame, "表2 Excel", self.file2_var, self.choose_file2)
        self._build_file_row(frame, "表数据量统计", self.table_stats_var, self.choose_table_stats_file)
        self._build_file_row(frame, "输出目录", self.output_dir_var, self.choose_output_dir, select_file=False)

        option_frame = ttk.Frame(frame)
        option_frame.pack(fill="x", pady=(18, 10))

        ttk.Label(
            option_frame,
            text="相同判断规则",
            font=("Microsoft YaHei UI", 10, "bold"),
        ).pack(anchor="w")

        ttk.Checkbutton(
            option_frame,
            text="合并判断相同SQL（忽略空白差异，并合并同一文件内重复SQL）",
            variable=self.relaxed_match_var,
        ).pack(anchor="w")

        tip = ttk.Label(
            option_frame,
            text="提示：不勾选时严格按原始 SQL 比较；勾选后会按更适合合并的规则判断相同 SQL。",
            foreground="#666666",
        )
        tip.pack(anchor="w", pady=(6, 0))

        threshold_row = ttk.Frame(option_frame)
        threshold_row.pack(fill="x", pady=(10, 0))
        ttk.Label(threshold_row, text="大表阈值(记录数)", width=16).pack(side="left")
        ttk.Entry(threshold_row, textvariable=self.large_table_threshold_var, width=18).pack(side="left")
        ttk.Label(
            threshold_row,
            text="默认 1000000；仅在上传“表数据量统计”时生效",
            foreground="#666666",
        ).pack(side="left", padx=(10, 0))

        button_frame = ttk.Frame(frame)
        button_frame.pack(fill="x", pady=(24, 12))

        ttk.Button(button_frame, text="开始比较", command=self.run_compare).pack(side="left")
        ttk.Button(button_frame, text="打开输出目录", command=self.open_output_dir).pack(side="left", padx=(12, 0))
        ttk.Button(button_frame, text="规则说明", command=self.show_rules).pack(side="left", padx=(12, 0))

        status_title = ttk.Label(frame, text="运行状态", font=("Microsoft YaHei UI", 10, "bold"))
        status_title.pack(anchor="w", pady=(8, 6))

        status_label = ttk.Label(
            frame,
            textvariable=self.status_var,
            relief="solid",
            padding=12,
            anchor="w",
            justify="left",
        )
        status_label.pack(fill="both", expand=True)

    def _build_file_row(
        self,
        parent: ttk.Frame,
        label_text: str,
        variable: tk.StringVar,
        command,
        select_file: bool = True,
    ) -> None:
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=8)

        ttk.Label(row, text=label_text, width=12).pack(side="left")
        ttk.Entry(row, textvariable=variable).pack(side="left", fill="x", expand=True, padx=(0, 10))
        button_text = "选择文件" if select_file else "选择目录"
        ttk.Button(row, text=button_text, command=command).pack(side="left")

    def choose_file1(self) -> None:
        path = self._select_excel_file()
        if path:
            self.file1_var.set(path)

    def choose_file2(self) -> None:
        path = self._select_excel_file()
        if path:
            self.file2_var.set(path)

    def choose_output_dir(self) -> None:
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_dir_var.set(path)

    def choose_table_stats_file(self) -> None:
        path = self._select_excel_file(title="选择表数据量统计文件")
        if path:
            self.table_stats_var.set(path)

    def _select_excel_file(self, title: str = "选择 Excel 文件") -> str:
        return filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )

    def run_compare(self) -> None:
        file1 = self.file1_var.get().strip()
        file2 = self.file2_var.get().strip()
        table_stats_file = self.table_stats_var.get().strip()
        output_dir = self.output_dir_var.get().strip()

        if not file1 or not file2:
            messagebox.showwarning("缺少文件", "请先选择表1和表2的 Excel 文件。")
            return

        if not output_dir:
            messagebox.showwarning("缺少输出目录", "请选择输出目录。")
            return

        try:
            large_table_threshold = int(self.large_table_threshold_var.get().strip())
        except Exception:
            messagebox.showwarning("阈值无效", "大表阈值请输入整数。")
            return

        if large_table_threshold < 0:
            messagebox.showwarning("阈值无效", "大表阈值不能小于 0。")
            return

        self.status_var.set("正在比较，请稍候...")
        self.root.update_idletasks()

        try:
            relaxed_match = self.relaxed_match_var.get()
            output_name = f"SQL比较结果_{Path(file1).stem}_VS_{Path(file2).stem}.xlsx"
            output_path = os.path.join(output_dir, output_name)
            if os.path.exists(output_path):
                should_overwrite = messagebox.askyesno(
                    "文件已存在",
                    f"结果文件已存在：\n{output_path}\n\n是否覆盖？",
                )
                if not should_overwrite:
                    self.status_var.set("已取消生成：存在同名结果文件，且未选择覆盖。")
                    return

            output_path, stats, _, _ = compare_sql_files(
                file1=file1,
                file2=file2,
                output_dir=output_dir,
                ignore_whitespace=relaxed_match,
                deduplicate_within_file=relaxed_match,
                table_stats_file=table_stats_file,
                large_table_threshold=large_table_threshold,
            )
        except BadZipFile:
            msg = (
                "文件扩展名看起来像 Excel，但文件内容不是标准的 .xlsx/.xlsm 格式。\n"
                "请确认不是把 CSV、截图导出文件或临时文件误当成 Excel 选进来了。"
            )
            self.status_var.set(f"比较失败：{msg}")
            messagebox.showerror("比较失败", msg)
            return
        except Exception as exc:
            self.status_var.set(f"比较失败：{exc}")
            traceback.print_exc()
            messagebox.showerror("比较失败", f"{exc}")
            return

        self.status_var.set(
            "比较完成。\n"
            f"结果文件：{output_path}\n\n"
            f"表1总数：{stats['表1总数']}\n"
            f"表2总数：{stats['表2总数']}\n"
            f"仅表1：{stats['仅表1']}\n"
            f"仅表2：{stats['仅表2']}\n"
            f"共有：{stats['共有']}\n"
            f"相似SQL类数：{stats['相似类数']}\n"
            f"合并判断相同SQL：{'是' if relaxed_match else '否'}\n"
            f"表内去重：{stats['表内去重']}\n"
            f"大表阈值：{stats['大表阈值']}\n\n"
            "生成内容：\n"
            f"1. 原表_{Path(file1).stem}\n"
            f"2. 原表_{Path(file2).stem}\n"
            f"3. 仅{Path(file1).stem}\n"
            f"4. 仅{Path(file2).stem}\n"
            f"5. 共有({Path(file1).stem})\n"
            f"6. 共有({Path(file2).stem})\n"
            "7. 相似SQL归类汇总\n"
            "8. 相似SQL明细"
        )
        messagebox.showinfo("比较完成", f"结果已生成：\n{output_path}")

    def open_output_dir(self) -> None:
        output_dir = self.output_dir_var.get().strip()
        if not output_dir:
            messagebox.showwarning("未设置目录", "请先选择输出目录。")
            return

        path = Path(output_dir)
        if not path.exists():
            messagebox.showwarning("目录不存在", "当前输出目录不存在，请重新选择。")
            return

        if sys.platform.startswith("win"):
            os.startfile(str(path))
        elif sys.platform == "darwin":
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}"')

    def show_rules(self) -> None:
        message = (
            "一、SQL 比较规则\n"
            "1. 以“SQL语句”列作为比较依据。\n"
            "2. 程序会自动在前 30 行里寻找真正表头，不要求第一行就是表头。\n"
            "3. 支持标准 .xlsx / .xlsm / .xls；对部分伪装成 .xls 的 HTML 表格也会尝试兼容读取。\n"
            "4. “合并判断相同SQL”勾选后，会同时忽略空白差异，并合并同一文件内重复 SQL。\n"
            "5. 不勾选时，严格按原始 SQL 比较，并保留表内重复行。\n\n"
            "6. 如果上传“表数据量统计”，程序会根据大表阈值判断相似 SQL 是否涉及大表。\n\n"
            "二、相似 SQL 归类规则\n"
            "1. 相似归类不是按业务语义，而是按标准化后的 SQL 骨架分组。\n"
            "2. 归类时会统一大小写和空白格式。\n"
            "3. 字符串常量会替换成 ?str?。\n"
            "4. 日期会替换成 ?date?，8 位日期串会替换成 ?date8?。\n"
            "5. 数字会替换成 ?num?。\n"
            "6. IN (...) 会折叠成 in(?list?)，VALUES (...) 会折叠成 values(?vals?)。\n"
            "7. 两条 SQL 标准化后完全一致，才会归到同一个“相似类ID”。\n\n"
            "三、汇总表字段说明\n"
            "1. 是否存在完全相同SQL：表示两个表里是否出现过完全一致的原始 SQL。\n"
            "2. 跨表原因：可能是“跨表完全相同 / 单表内完全相同 / 仅日期不同 / IN列表不同 / 仅参数不同 / 结构相似 / 仅单表出现”。\n"
            "3. 对应表内第几条：显示这一类 SQL 分别出现在各原表里的第几条，方便回原表定位。\n"
            "4. 规则判断是否有大表 / 规则判断涉及到的大表名称：基于“表数据量统计”文件中的表名和记录数判断。"
        )
        messagebox.showinfo("规则说明", message)


def main() -> None:
    root = tk.Tk()
    style = ttk.Style()
    if "vista" in style.theme_names():
        style.theme_use("vista")

    SqlDiffApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
