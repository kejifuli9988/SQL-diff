import os
import re
import sys
import traceback
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, Tuple

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


SQL_COLUMN_NAME = "SQL语句"


def normalize_sql(value: object, ignore_whitespace: bool = True) -> str:
    if pd.isna(value):
        return ""

    text = str(value).replace("\u3000", " ").strip()
    if ignore_whitespace:
        text = re.sub(r"\s+", " ", text)
    return text


def load_excel(file_path: str) -> pd.DataFrame:
    suffix = Path(file_path).suffix.lower()
    if suffix not in {".xlsx", ".xls"}:
        raise ValueError(f"不支持的文件类型: {suffix}")

    return pd.read_excel(file_path)


def ensure_sql_column(df: pd.DataFrame, file_label: str) -> None:
    if SQL_COLUMN_NAME not in df.columns:
        raise ValueError(f"{file_label} 中未找到表头 “{SQL_COLUMN_NAME}”")


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


def compare_sql_files(
    file1: str,
    file2: str,
    output_dir: str,
    ignore_whitespace: bool = True,
) -> str:
    os.makedirs(output_dir, exist_ok=True)

    df1 = load_excel(file1)
    df2 = load_excel(file2)

    ensure_sql_column(df1, "表1")
    ensure_sql_column(df2, "表2")

    map1, order1 = build_sql_map(df1, ignore_whitespace)
    map2, order2 = build_sql_map(df2, ignore_whitespace)

    set1 = set(map1.keys())
    set2 = set(map2.keys())

    only1_keys = [key for key in order1 if key in (set1 - set2)]
    only2_keys = [key for key in order2 if key in (set2 - set1)]
    both1_keys = [key for key in order1 if key in (set1 & set2)]
    both2_keys = [key for key in order2 if key in (set1 & set2)]

    only1_df = build_dataframe_from_rows((map1[key] for key in only1_keys), df1.columns)
    only2_df = build_dataframe_from_rows((map2[key] for key in only2_keys), df2.columns)
    both1_df = build_dataframe_from_rows((map1[key] for key in both1_keys), df1.columns)
    both2_df = build_dataframe_from_rows((map2[key] for key in both2_keys), df2.columns)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = os.path.join(output_dir, f"SQL比较结果_{timestamp}.xlsx")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        only1_df.to_excel(writer, sheet_name="表1不在表2", index=False)
        only2_df.to_excel(writer, sheet_name="表2不在表1", index=False)
        both1_df.to_excel(writer, sheet_name="两个表都有_表1格式", index=False)
        both2_df.to_excel(writer, sheet_name="两个表都有_表2格式", index=False)

    return output_path


class SqlDiffApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("SQL 语句 Excel 比较工具")
        self.root.geometry("720x420")
        self.root.minsize(680, 400)

        self.file1_var = tk.StringVar()
        self.file2_var = tk.StringVar()
        self.output_dir_var = tk.StringVar(value=str(Path.home() / "Desktop"))
        self.ignore_whitespace_var = tk.BooleanVar(value=True)
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
        self._build_file_row(frame, "输出目录", self.output_dir_var, self.choose_output_dir, select_file=False)

        option_frame = ttk.Frame(frame)
        option_frame.pack(fill="x", pady=(18, 10))

        ttk.Checkbutton(
            option_frame,
            text="忽略空白差异（换行、多个空格、首尾空格）",
            variable=self.ignore_whitespace_var,
        ).pack(anchor="w")

        tip = ttk.Label(
            option_frame,
            text="提示：如果同一个文件中同一条 SQL 重复出现，程序会保留首次出现的那一行。",
            foreground="#666666",
        )
        tip.pack(anchor="w", pady=(6, 0))

        button_frame = ttk.Frame(frame)
        button_frame.pack(fill="x", pady=(24, 12))

        ttk.Button(button_frame, text="开始比较", command=self.run_compare).pack(side="left")
        ttk.Button(button_frame, text="打开输出目录", command=self.open_output_dir).pack(side="left", padx=(12, 0))

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

    def _select_excel_file(self) -> str:
        return filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )

    def run_compare(self) -> None:
        file1 = self.file1_var.get().strip()
        file2 = self.file2_var.get().strip()
        output_dir = self.output_dir_var.get().strip()

        if not file1 or not file2:
            messagebox.showwarning("缺少文件", "请先选择表1和表2的 Excel 文件。")
            return

        if not output_dir:
            messagebox.showwarning("缺少输出目录", "请选择输出目录。")
            return

        self.status_var.set("正在比较，请稍候...")
        self.root.update_idletasks()

        try:
            output_path = compare_sql_files(
                file1=file1,
                file2=file2,
                output_dir=output_dir,
                ignore_whitespace=self.ignore_whitespace_var.get(),
            )
        except Exception as exc:
            self.status_var.set(f"比较失败：{exc}")
            traceback.print_exc()
            messagebox.showerror("比较失败", f"{exc}")
            return

        self.status_var.set(
            "比较完成。\n"
            f"结果文件：{output_path}\n\n"
            "生成内容：\n"
            "1. 表1不在表2\n"
            "2. 表2不在表1\n"
            "3. 两个表都有_表1格式\n"
            "4. 两个表都有_表2格式"
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


def main() -> None:
    root = tk.Tk()
    style = ttk.Style()
    if "vista" in style.theme_names():
        style.theme_use("vista")

    SqlDiffApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
