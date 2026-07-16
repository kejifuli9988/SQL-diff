import os
import re
import sys
import traceback
import hashlib
import json
from difflib import SequenceMatcher
from zipfile import BadZipFile
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set, Tuple

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


SQL_COLUMN_NAME = "SQL语句"
TABLE_STATS_NAME_COLUMN = "表名"
TABLE_STATS_ROWS_COLUMN = "记录数"
COMPARE_MODE_STRICT = "strict"
COMPARE_MODE_SMART = "smart"
DEFAULT_HISTORY_ENRICH_COLUMNS = [
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
SETTINGS_FILE = Path(__file__).with_name("sql_diff_gui_settings.json")


def load_enrich_columns_config() -> List[str]:
    if SETTINGS_FILE.exists():
        try:
            data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            columns = data.get("history_enrich_columns", [])
            if isinstance(columns, list):
                cleaned = []
                seen: Set[str] = set()
                for item in columns:
                    text = str(item).strip()
                    if text and text not in seen:
                        seen.add(text)
                        cleaned.append(text)
                if cleaned:
                    return cleaned
        except Exception:
            pass
    return DEFAULT_HISTORY_ENRICH_COLUMNS.copy()


def save_enrich_columns_config(columns: List[str]) -> None:
    SETTINGS_FILE.write_text(
        json.dumps({"history_enrich_columns": columns}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


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


def build_similarity_signature(sql: str, ignore_whitespace: bool = True) -> str:
    s = normalize_sql(sql, ignore_whitespace=ignore_whitespace).lower()
    s = re.sub(r"'(?:''|[^'])*'", "?str?", s)
    s = re.sub(r"\b\d{4}[-/]\d{1,2}[-/]\d{1,2}(?: \d{1,2}:\d{1,2}:\d{1,2})?\b", "?date?", s)
    s = re.sub(r"\b\d{8}\b", "?date8?", s)
    s = re.sub(r"\b\d+\b", "?num?", s)
    s = re.sub(r"\bin\s*\((?:[^()]*?)\)", "in(?list?)", s)
    s = re.sub(r"\bvalues\s*\((?:[^()]*?)\)", "values(?vals?)", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def build_date_only_signature(sql: str, ignore_whitespace: bool = True) -> str:
    s = normalize_sql(sql, ignore_whitespace=ignore_whitespace).lower()
    s = re.sub(r"\b\d{4}[-/]\d{1,2}[-/]\d{1,2}(?: \d{1,2}:\d{1,2}:\d{1,2})?\b", "?date?", s)
    s = re.sub(r"\b\d{8}\b", "?date8?", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def build_inlist_signature(sql: str, ignore_whitespace: bool = True) -> str:
    s = normalize_sql(sql, ignore_whitespace=ignore_whitespace).lower()
    s = re.sub(r"\bin\s*\((?:[^()]*?)\)", "in(?list?)", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def build_parameter_signature(sql: str, ignore_whitespace: bool = True) -> str:
    s = normalize_sql(sql, ignore_whitespace=ignore_whitespace).lower()
    s = re.sub(r"'(?:''|[^'])*'", "?str?", s)
    s = re.sub(r"\b\d{4}[-/]\d{1,2}[-/]\d{1,2}(?: \d{1,2}:\d{1,2}:\d{1,2})?\b", "?date?", s)
    s = re.sub(r"\b\d{8}\b", "?date8?", s)
    s = re.sub(r"\b\d+\b", "?num?", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def build_paging_signature(sql: str, ignore_whitespace: bool = True) -> str:
    s = normalize_sql(sql, ignore_whitespace=ignore_whitespace).lower()
    s = re.sub(r"rownum_\s*>\s*\d+", "rownum_>?page?", s)
    s = re.sub(r"rownum_\s*<=\s*\d+", "rownum_<=?page?", s)
    s = re.sub(r"\blimit\s+\d+\b", "limit ?page?", s)
    s = re.sub(r"\boffset\s+\d+\b", "offset ?page?", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _find_top_level_keyword(sql: str, keyword: str) -> int:
    pattern = keyword.lower()
    text = sql.lower()
    depth = 0
    i = 0
    while i < len(text):
        ch = text[i]
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth = max(0, depth - 1)
        elif depth == 0 and text.startswith(pattern, i):
            prev_ok = i == 0 or not text[i - 1].isalnum()
            end = i + len(pattern)
            next_ok = end >= len(text) or not text[end].isalnum()
            if prev_ok and next_ok:
                return i
        i += 1
    return -1


def build_select_body_signature(sql: str, ignore_whitespace: bool = True) -> str:
    s = normalize_sql(sql, ignore_whitespace=ignore_whitespace).lower()
    if not s.startswith("select "):
        return s
    from_idx = _find_top_level_keyword(s, " from ")
    if from_idx == -1:
        return s
    body = s[from_idx:]
    body = re.sub(r"'(?:''|[^'])*'", "?str?", body)
    body = re.sub(r"\b\d{4}[-/]\d{1,2}[-/]\d{1,2}(?: \d{1,2}:\d{1,2}:\d{1,2})?\b", "?date?", body)
    body = re.sub(r"\b\d{8}\b", "?date8?", body)
    body = re.sub(r"\b\d+\b", "?num?", body)
    body = re.sub(r"\s+", " ", body).strip()
    return body


def build_where_base_signature(sql: str, ignore_whitespace: bool = True) -> str:
    s = normalize_sql(sql, ignore_whitespace=ignore_whitespace).lower()
    where_idx = _find_top_level_keyword(s, " where ")
    if where_idx == -1:
        return s
    suffix_starts = []
    for keyword in [" group by ", " order by ", " union ", " union all ", " having "]:
        idx = _find_top_level_keyword(s, keyword)
        if idx > where_idx:
            suffix_starts.append(idx)
    tail_idx = min(suffix_starts) if suffix_starts else len(s)
    prefix = s[:where_idx]
    suffix = s[tail_idx:]
    result = prefix + " where ?conds? " + suffix
    result = re.sub(r"\s+", " ", result).strip()
    return result


def _build_row_signatures(sql: str, ignore_whitespace: bool) -> Dict[str, str]:
    return {
        "exact": normalize_sql(sql, ignore_whitespace=ignore_whitespace),
        "parameter": build_parameter_signature(sql, ignore_whitespace=ignore_whitespace),
        "date": build_date_only_signature(sql, ignore_whitespace=ignore_whitespace),
        "inlist": build_inlist_signature(sql, ignore_whitespace=ignore_whitespace),
        "paging": build_paging_signature(sql, ignore_whitespace=ignore_whitespace),
        "select_body": build_select_body_signature(sql, ignore_whitespace=ignore_whitespace),
        "where_base": build_where_base_signature(sql, ignore_whitespace=ignore_whitespace),
        "similarity": build_similarity_signature(sql, ignore_whitespace=ignore_whitespace),
    }


def _build_match_info(
    df_self: pd.DataFrame,
    df_other: pd.DataFrame,
    ignore_whitespace: bool,
    compare_mode: str,
) -> Dict[int, Dict[str, object]]:
    other_infos = []
    other_exact_positions: Dict[str, List[int]] = {}
    for _, row in df_other.reset_index(drop=True).iterrows():
        sql = str(row.get(SQL_COLUMN_NAME, ""))
        sigs = _build_row_signatures(sql, ignore_whitespace=ignore_whitespace)
        source_seq = int(row.get("_source_seq", 0))
        other_infos.append({"seq": source_seq, "sql": sql, "sigs": sigs})
        other_exact_positions.setdefault(sigs["exact"], []).append(source_seq)

    result: Dict[int, Dict[str, object]] = {}
    for _, row in df_self.reset_index(drop=True).iterrows():
        sql = str(row.get(SQL_COLUMN_NAME, ""))
        source_seq = int(row.get("_source_seq", 0))
        sigs = _build_row_signatures(sql, ignore_whitespace=ignore_whitespace)

        exact_positions = other_exact_positions.get(sigs["exact"], [])
        if exact_positions:
            result[source_seq] = {
                "matched": True,
                "match_type": "完全相同",
                "match_reason": "完全相同",
                "other_positions": "、".join(str(x) for x in exact_positions),
            }
            continue

        if compare_mode != COMPARE_MODE_SMART:
            result[source_seq] = {
                "matched": False,
                "match_type": "未匹配",
                "match_reason": "未找到完全相同SQL",
                "other_positions": "",
            }
            continue

        best_match: Optional[Dict[str, object]] = None
        best_score = 0.0
        for other in other_infos:
            other_sigs = other["sigs"]
            reason = ""
            if sigs["parameter"] == other_sigs["parameter"]:
                reason = "仅参数不同"
            elif sigs["date"] == other_sigs["date"]:
                reason = "仅日期不同"
            elif sigs["inlist"] == other_sigs["inlist"]:
                reason = "IN列表不同"
            elif sigs["paging"] == other_sigs["paging"] and sigs["exact"] != other_sigs["exact"]:
                reason = "分页条件不同"
            elif sigs["select_body"] == other_sigs["select_body"] and sigs["exact"] != other_sigs["exact"]:
                reason = "SELECT字段增加/减少"
            elif sigs["where_base"] == other_sigs["where_base"] and sigs["exact"] != other_sigs["exact"]:
                reason = "WHERE条件增加/减少"

            if reason:
                best_match = {
                    "matched": True,
                    "match_type": "智能匹配",
                    "match_reason": reason,
                    "other_positions": str(other["seq"]),
                }
                break

            score = SequenceMatcher(None, sigs["similarity"], other_sigs["similarity"]).ratio()
            if score >= 0.80 and score > best_score:
                best_score = score
                best_match = {
                    "matched": True,
                    "match_type": "智能匹配",
                    "match_reason": "高相似度匹配",
                    "other_positions": str(other["seq"]),
                }

        if best_match is not None:
            result[source_seq] = best_match
        else:
            result[source_seq] = {
                "matched": False,
                "match_type": "未匹配",
                "match_reason": "未找到符合规则的SQL",
                "other_positions": "",
            }
    return result


def _attach_match_columns(
    df: pd.DataFrame,
    match_info: Dict[int, Dict[str, object]],
    matched: bool,
) -> pd.DataFrame:
    output = df.copy()
    match_types = []
    match_reasons = []
    other_positions = []
    keep_mask = []
    for _, row in output.iterrows():
        source_seq = int(row.get("_source_seq", 0))
        info = match_info.get(source_seq, {
            "matched": False,
            "match_type": "未匹配",
            "match_reason": "未找到匹配SQL",
            "other_positions": "",
        })
        keep_mask.append(bool(info["matched"]) == matched)
        match_types.append(str(info["match_type"]))
        match_reasons.append(str(info["match_reason"]))
        other_positions.append(str(info["other_positions"]))
    output["匹配方式"] = match_types
    output["匹配原因"] = match_reasons
    output["对方表第几条"] = other_positions
    output["组内匹配条目"] = ""
    output = output[pd.Series(keep_mask, index=output.index)].copy()
    return output


def _apply_single_file_match_labels(
    df: pd.DataFrame,
    source_file_name: str,
    single_file_reason_map: Dict[Tuple[str, int], str],
    single_file_group_map: Dict[Tuple[str, int], str],
) -> pd.DataFrame:
    if df.empty or "_source_seq" not in df.columns:
        return df

    output = df.copy()
    for col in ["匹配方式", "匹配原因", "对方表第几条", "组内匹配条目"]:
        if col not in output.columns:
            output[col] = ""
    for idx, row in output.iterrows():
        key = (source_file_name, int(row.get("_source_seq", 0)))
        reason = single_file_reason_map.get(key)
        if reason:
            output.at[idx, "匹配方式"] = "单表内匹配"
            output.at[idx, "匹配原因"] = reason
            output.at[idx, "对方表第几条"] = ""
            output.at[idx, "组内匹配条目"] = single_file_group_map.get(key, "")
    return output


def _build_single_file_maps_from_summary(
    summary_df: pd.DataFrame,
    detail_df: pd.DataFrame,
) -> Tuple[Dict[Tuple[str, int], str], Dict[Tuple[str, int], str]]:
    reason_map: Dict[Tuple[str, int], str] = {}
    group_map: Dict[Tuple[str, int], str] = {}
    if summary_df.empty or detail_df.empty:
        return reason_map, group_map

    required_cols = {"相似类ID", "涉及文件数", "重复条数"}
    if not required_cols.issubset(set(summary_df.columns)):
        return reason_map, group_map

    single_groups = summary_df[
        (pd.to_numeric(summary_df["涉及文件数"], errors="coerce").fillna(0).astype(int) == 1)
        & (pd.to_numeric(summary_df["重复条数"], errors="coerce").fillna(0).astype(int) > 1)
    ].copy()
    if single_groups.empty:
        return reason_map, group_map

    group_ids = set(single_groups["相似类ID"].astype(str).tolist())
    for group_id in group_ids:
        grp = detail_df[detail_df["相似类ID"].astype(str) == group_id]
        if grp.empty:
            continue
        reason = infer_single_file_match_reason(grp)
        grouped_by_file: Dict[str, List[int]] = {}
        for _, row in grp.iterrows():
            file_name = str(row["来源文件"])
            seq = int(row["对应表内第几条"])
            grouped_by_file.setdefault(file_name, []).append(seq)

        text_parts = []
        for file_name, seqs in grouped_by_file.items():
            seq_text = "、".join(str(x) for x in sorted(seqs))
            text_parts.append(f"{file_name}: {seq_text}")
        summary_text = " | ".join(text_parts)

        for file_name, seqs in grouped_by_file.items():
            for seq in seqs:
                key = (file_name, seq)
                reason_map[key] = reason
                group_map[key] = summary_text
    return reason_map, group_map


def _collapse_single_file_match_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "匹配方式" not in df.columns or "组内匹配条目" not in df.columns:
        return df

    output = df.copy()
    mask = (
        output["匹配方式"].astype(str).eq("单表内匹配")
        & output["组内匹配条目"].astype(str).str.strip().ne("")
    )
    if not mask.any():
        return output

    grouped = output[mask].drop_duplicates(subset=["组内匹配条目"], keep="first")
    others = output[~mask]
    output = pd.concat([others, grouped], axis=0).sort_index()
    return output


def _drop_only_sheet_helper_columns(df: pd.DataFrame) -> pd.DataFrame:
    output = df.copy()
    removable = [col for col in ["对方表第几条"] if col in output.columns]
    if removable:
        output = output.drop(columns=removable)
    return output


def _hide_internal_export_columns(df: pd.DataFrame) -> pd.DataFrame:
    output = df.copy()
    hidden_columns = [col for col in ["相似类ID"] if col in output.columns]
    if hidden_columns:
        output = output.drop(columns=hidden_columns)
    return output


def _finalize_shared_sheet_columns(df: pd.DataFrame) -> pd.DataFrame:
    output = df.copy()
    if "_source_seq" in output.columns:
        insert_at = len(output.columns)
        if "序号" in output.columns:
            insert_at = output.columns.get_loc("序号") + 1
        elif "匹配方式" in output.columns:
            insert_at = output.columns.get_loc("匹配方式")
        output.insert(insert_at, "对应表内第几条", output["_source_seq"].astype(int))
    return output


def _parse_position_text(position_text: object) -> Dict[str, List[int]]:
    result: Dict[str, List[int]] = {}
    text = str(position_text)
    for match in re.finditer(r"([^:；|]+?):\s*([0-9、]+)", text):
        file_name = match.group(1).strip()
        if "：" in file_name:
            file_name = file_name.split("：")[-1].strip()
        seqs = [int(x) for x in match.group(2).split("、") if x.strip().isdigit()]
        if file_name and seqs:
            result.setdefault(file_name, []).extend(seqs)
    return result


def _apply_group_summary_to_shared_sheet(
    df: pd.DataFrame,
    source_file_name: str,
    summary_df: pd.DataFrame,
    detail_df: pd.DataFrame,
) -> pd.DataFrame:
    if (
        df.empty
        or "_source_seq" not in df.columns
        or "相似类ID" not in summary_df.columns
        or "对应表内第几条" not in summary_df.columns
        or "相似类ID" not in detail_df.columns
    ):
        return df

    shared_summary = summary_df[
        pd.to_numeric(summary_df.get("涉及文件数"), errors="coerce").fillna(0).astype(int) >= 2
    ].copy()
    if shared_summary.empty:
        return df

    summary_text_by_seq: Dict[int, str] = {}
    for _, row in shared_summary.iterrows():
        summary_text = str(row["对应表内第几条"])
        positions = _parse_position_text(summary_text)
        for seq in positions.get(source_file_name, []):
            summary_text_by_seq[seq] = summary_text

    if not summary_text_by_seq:
        return df

    output = df.copy()
    for idx, row in output.iterrows():
        source_seq = int(row.get("_source_seq", 0))
        summary_text = summary_text_by_seq.get(source_seq)
        if not summary_text:
            continue
        output.at[idx, "对方表第几条"] = summary_text
    return output


def _collapse_shared_sheet_rows(
    df: pd.DataFrame,
    source_file_name: str,
    summary_df: pd.DataFrame,
    detail_df: pd.DataFrame,
) -> pd.DataFrame:
    if df.empty or "_source_seq" not in df.columns or "相似类ID" not in detail_df.columns:
        return df

    shared_group_ids = set(
        summary_df.loc[
            pd.to_numeric(summary_df.get("涉及文件数"), errors="coerce").fillna(0).astype(int) >= 2,
            "相似类ID",
        ].astype(str).tolist()
    )
    if not shared_group_ids:
        return df

    all_shared_seqs: Set[int] = set()
    keep_shared_seqs: Set[int] = set()
    for _, row in summary_df[
        summary_df["相似类ID"].astype(str).isin(shared_group_ids)
    ].iterrows():
        positions = _parse_position_text(row["对应表内第几条"])
        seqs = sorted(set(positions.get(source_file_name, [])))
        if not seqs:
            continue
        all_shared_seqs.update(seqs)
        keep_shared_seqs.add(seqs[0])

    if not all_shared_seqs:
        return df

    output = df.copy()
    is_shared_row = output["_source_seq"].astype(int).isin(all_shared_seqs)
    keep_mask = (~is_shared_row) | output["_source_seq"].astype(int).isin(keep_shared_seqs)
    output = output[keep_mask].copy()
    return output


def _move_single_file_summary_rows_from_both_to_only(
    only_df: pd.DataFrame,
    both_df: pd.DataFrame,
    source_file_name: str,
    summary_df: pd.DataFrame,
    detail_df: pd.DataFrame,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if (
        only_df.empty and both_df.empty
    ) or summary_df.empty or detail_df.empty or "_source_seq" not in both_df.columns:
        return only_df, both_df

    single_summary = summary_df.loc[
        summary_df["匹配方式"].astype(str) == "单表内匹配"
    ].copy()
    if single_summary.empty:
        return only_df, both_df

    # Only move groups that truly belong to one file. If the displayed group text
    # already mentions both files, keep it in shared sheets.
    single_summary = single_summary[
        (single_summary["来源文件"].astype(str) == source_file_name)
        & (pd.to_numeric(single_summary["涉及文件数"], errors="coerce").fillna(0).astype(int) == 1)
    ].copy()
    single_group_ids = set(single_summary["相似类ID"].astype(str).tolist())
    if not single_group_ids:
        return only_df, both_df

    seqs_to_move = set(
        detail_df.loc[
            (detail_df["来源文件"].astype(str) == source_file_name)
            & (detail_df["相似类ID"].astype(str).isin(single_group_ids)),
            "对应表内第几条",
        ].astype(int).tolist()
    )
    if not seqs_to_move:
        return only_df, both_df

    move_mask = both_df["_source_seq"].astype(int).isin(seqs_to_move)
    if not move_mask.any():
        return only_df, both_df

    moved_rows = both_df[move_mask].copy()
    kept_rows = both_df[~move_mask].copy()
    only_df = pd.concat([only_df, moved_rows], axis=0).sort_index()
    return only_df, kept_rows


def _move_strict_single_side_rows_from_both_to_only(
    only_df: pd.DataFrame,
    both_df: pd.DataFrame,
    source_file_name: str,
    summary_df: pd.DataFrame,
    detail_df: pd.DataFrame,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if (
        only_df.empty and both_df.empty
    ) or summary_df.empty or detail_df.empty or "_source_seq" not in both_df.columns:
        return only_df, both_df

    if "跨表原因" not in summary_df.columns:
        return only_df, both_df

    target_group_ids = set(
        summary_df.loc[
            summary_df["跨表原因"].astype(str) == "仅单表出现",
            "相似类ID",
        ].astype(str).tolist()
    )
    if not target_group_ids:
        return only_df, both_df

    seqs_to_move = set(
        detail_df.loc[
            (detail_df["来源文件"].astype(str) == source_file_name)
            & (detail_df["相似类ID"].astype(str).isin(target_group_ids)),
            "对应表内第几条",
        ].astype(int).tolist()
    )
    if not seqs_to_move:
        return only_df, both_df

    move_mask = both_df["_source_seq"].astype(int).isin(seqs_to_move)
    if not move_mask.any():
        return only_df, both_df

    moved_rows = both_df[move_mask].copy()
    kept_rows = both_df[~move_mask].copy()
    only_df = pd.concat([only_df, moved_rows], axis=0).sort_index()
    return only_df, kept_rows


def _sync_shared_summary_rows(
    only_df: pd.DataFrame,
    both_df: pd.DataFrame,
    source_file_name: str,
    summary_df: pd.DataFrame,
    detail_df: pd.DataFrame,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if summary_df.empty or detail_df.empty or "_source_seq" not in only_df.columns:
        return only_df, both_df

    shared_group_ids = set(
        summary_df.loc[
            pd.to_numeric(summary_df.get("涉及文件数"), errors="coerce").fillna(0).astype(int) >= 2,
            "相似类ID",
        ].astype(str).tolist()
    )
    if not shared_group_ids:
        return only_df, both_df

    shared_seqs = set(
        detail_df.loc[
            (detail_df["来源文件"].astype(str) == source_file_name)
            & (detail_df["相似类ID"].astype(str).isin(shared_group_ids)),
            "对应表内第几条",
        ].astype(int).tolist()
    )
    if not shared_seqs:
        return only_df, both_df

    move_mask = only_df["_source_seq"].astype(int).isin(shared_seqs)
    if move_mask.any():
        moved_rows = only_df[move_mask].copy()
        both_df = pd.concat([both_df, moved_rows], axis=0).sort_index()
        only_df = only_df[~move_mask].copy()
    return only_df, both_df


def build_smart_group_id(
    current_file_name: str,
    current_seq: int,
    other_file_name: str,
    other_positions: str,
) -> str:
    tokens = [f"{current_file_name}#{current_seq}"]
    for part in str(other_positions).replace("|", "、").split("、"):
        value = part.strip()
        if value:
            tokens.append(f"{other_file_name}#{value}")
    canonical = "|".join(sorted(set(tokens)))
    return "M" + hashlib.md5(canonical.encode("utf-8")).hexdigest()[:8].upper()


def build_smart_group_ids_from_detail(detail_df: pd.DataFrame) -> Dict[str, str]:
    adjacency: Dict[str, Set[str]] = {}

    def ensure_node(node: str) -> None:
        adjacency.setdefault(node, set())

    for _, row in detail_df.iterrows():
        file_name = str(row.get("来源文件", ""))
        source_seq = str(row.get("对应表内第几条", ""))
        if not file_name or not source_seq:
            continue
        node = f"{file_name}#{source_seq}"
        ensure_node(node)
        other_file_candidates = [name for name in detail_df["来源文件"].drop_duplicates().tolist() if name != file_name]
        other_file_name = other_file_candidates[0] if other_file_candidates else ""
        if not other_file_name:
            continue
        for part in str(row.get("对方表第几条", "")).replace("|", "、").split("、"):
            value = part.strip()
            if not value:
                continue
            other_node = f"{other_file_name}#{value}"
            ensure_node(other_node)
            adjacency[node].add(other_node)
            adjacency[other_node].add(node)

    group_map: Dict[str, str] = {}
    visited: Set[str] = set()
    for node in adjacency:
        if node in visited or not adjacency[node]:
            continue
        stack = [node]
        component = []
        visited.add(node)
        while stack:
            current = stack.pop()
            component.append(current)
            for neighbor in adjacency[current]:
                if neighbor not in visited:
                    visited.add(neighbor)
                    stack.append(neighbor)
        group_id = "M" + hashlib.md5("|".join(sorted(component)).encode("utf-8")).hexdigest()[:8].upper()
        for item in component:
            group_map[item] = group_id
    return group_map


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


def build_position_summary(grp: pd.DataFrame, compare_mode: str) -> str:
    if compare_mode != COMPARE_MODE_SMART or "匹配原因" not in grp.columns:
        positions = []
        for file_name, sub in grp.groupby("来源文件", sort=False):
            seqs = "、".join(str(x) for x in sub["对应表内第几条"].tolist())
            positions.append(f"{file_name}: {seqs}")
        return " | ".join(positions)

    if grp["来源文件"].nunique() == 1 and len(grp) > 1:
        positions = []
        for file_name, sub in grp.groupby("来源文件", sort=False):
            seqs = "、".join(str(x) for x in sub["对应表内第几条"].tolist())
            positions.append(f"{file_name}: {seqs}")
        return f"单表内匹配：{' | '.join(positions)}"

    sections = []
    temp = grp.copy()
    temp["_summary_reason"] = temp["匹配原因"].astype(str)
    if "匹配方式" in temp.columns:
        temp.loc[temp["匹配方式"].astype(str) == "完全相同", "_summary_reason"] = "完全相同"

    for reason, reason_grp in temp.groupby("_summary_reason", sort=False):
        reason_positions = []
        for file_name, sub in reason_grp.groupby("来源文件", sort=False):
            seqs = "、".join(str(x) for x in sub["对应表内第几条"].tolist())
            reason_positions.append(f"{file_name}: {seqs}")
        sections.append(f"{reason}：{' | '.join(reason_positions)}")
    return "；".join(sections)


def infer_single_file_match_reason(grp: pd.DataFrame) -> str:
    sqls = grp["SQL语句"].astype(str).tolist()
    if len(set(sqls)) < len(sqls):
        return "完全相同"

    parameter_signatures = set(grp["参数归一特征"].astype(str).tolist()) if "参数归一特征" in grp.columns else set()
    date_signatures = set(grp["日期归一特征"].astype(str).tolist()) if "日期归一特征" in grp.columns else set()
    inlist_signatures = set(grp["IN归一特征"].astype(str).tolist()) if "IN归一特征" in grp.columns else set()

    if len(parameter_signatures) == 1:
        return "仅参数不同"
    if len(date_signatures) == 1:
        return "仅日期不同"
    if len(inlist_signatures) == 1:
        return "IN列表不同"

    paging_signatures = {build_paging_signature(sql, ignore_whitespace=True) for sql in sqls}
    if len(paging_signatures) == 1:
        return "分页条件不同"

    select_body_signatures = {build_select_body_signature(sql, ignore_whitespace=True) for sql in sqls}
    if len(select_body_signatures) == 1:
        return "SELECT字段增加/减少"

    where_base_signatures = {build_where_base_signature(sql, ignore_whitespace=True) for sql in sqls}
    if len(where_base_signatures) == 1:
        return "WHERE条件增加/减少"

    return "高相似度匹配"


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


def infer_performance_cause(
    sql: str,
    sql_type: str,
    matched_tables: List[str],
    large_table_map: Optional[Dict[str, int]] = None,
) -> Tuple[str, str, str]:
    normalized = normalize_sql(sql, ignore_whitespace=True).lower()
    join_count = len(re.findall(r"\bjoin\b", normalized))
    in_items = re.search(r"\bin\s*\(([^()]*)\)", normalized)
    in_count = 0
    if in_items:
        content = in_items.group(1).strip()
        if content:
            in_count = len([x for x in content.split(",") if x.strip()])

    has_between_date = bool(
        re.search(r"\bbetween\b.*?(?:\d{4}[-/]\d{1,2}[-/]\d{1,2}|\?date\?)", normalized)
    )
    has_order_group = any(token in normalized for token in [" order by ", " group by ", " distinct "])
    has_select_all = normalized.startswith("select *") or " select * " in normalized
    has_exists = " exists (" in normalized or " not exists (" in normalized
    has_union = " union " in normalized
    has_paging = "rownum" in normalized or " limit " in normalized or " offset " in normalized
    has_large_tables = bool(matched_tables)
    has_count = normalized.startswith("select count(")
    where_clause = normalized.split(" where ", 1)[1] if " where " in normalized else ""
    equality_count = len(re.findall(r"[a-zA-Z0-9_.$]+\s*=\s*(?:\?|:\w+|'[^']*'|\d+)", where_clause))

    if has_large_tables and sql_type in {"普通SELECT", "COUNT查询", "WITH查询"} and equality_count <= 1:
        return (
            "大表扫描",
            f"命中大表：{'、'.join(matched_tables)}；过滤条件较弱",
            "优先检查索引、过滤条件和是否能缩小扫描范围",
        )
    if has_paging:
        return (
            "分页过深",
            "命中分页语法（rownum/limit/offset）",
            "考虑改为基于索引的游标/主键翻页，避免深分页",
        )
    if in_count >= 10:
        return (
            "IN列表过长",
            f"检测到 IN 列表项较多（约 {in_count} 项）",
            "考虑临时表/批量表关联，或拆分请求减少 IN 列表长度",
        )
    if has_between_date:
        return (
            "时间范围过大",
            "检测到时间区间过滤",
            "检查是否能缩短时间窗口，或按时间字段建立更合适索引",
        )
    if join_count >= 3 or has_exists or has_union:
        return (
            "多表关联复杂",
            f"检测到 {join_count} 个 JOIN，或存在 EXISTS/UNION",
            "优先检查关联顺序、驱动表、索引和是否可拆分查询",
        )
    if has_order_group or has_count:
        return (
            "排序或聚合代价高",
            "检测到 ORDER BY / GROUP BY / DISTINCT / COUNT",
            "检查排序字段、分组字段索引，评估是否可减少聚合范围",
        )
    if sql_type in {"INSERT", "UPDATE", "DELETE"}:
        return (
            "写入或更新代价高",
            "属于 INSERT / UPDATE / DELETE 语句",
            "检查更新条件索引、锁竞争和批量写入方式",
        )
    if has_select_all:
        return (
            "返回列过多",
            "检测到 SELECT * 或宽字段查询",
            "只返回必要字段，避免大对象或宽表整行拉取",
        )
    return (
        "规则暂未明确归类",
        "未命中当前主要慢SQL规则",
        "建议结合执行计划、索引和表统计信息进一步分析",
    )


def build_summary_row(
    grp: pd.DataFrame,
    class_id: str,
    cross_reason: str,
    has_exact_match: str,
    large_table_map: Optional[Dict[str, int]] = None,
    enrich_columns: Optional[List[str]] = None,
    include_performance_rules: bool = False,
    compare_mode: str = COMPARE_MODE_STRICT,
) -> Dict[str, object]:
    matched_tables: List[str] = []
    if large_table_map:
        seen: Set[str] = set()
        for sql in grp["SQL语句"].astype(str):
            for table_name in extract_table_names(sql):
                if table_name in large_table_map and table_name not in seen:
                    seen.add(table_name)
                    matched_tables.append(table_name)

    summary_row = {
        "相似类ID": class_id,
        "SQL类型": grp["SQL类型"].iloc[0],
        "重复条数": len(grp),
        "涉及文件数": grp["来源文件"].nunique(),
        "是否存在完全相同SQL": has_exact_match,
    }
    if compare_mode == COMPARE_MODE_SMART:
        if grp["来源文件"].nunique() == 1 and len(grp) > 1:
            single_file_reason = infer_single_file_match_reason(grp)
            summary_row["匹配方式"] = "单表内匹配"
            summary_row["匹配原因"] = single_file_reason
        else:
            match_types = grp["匹配方式"].dropna().astype(str).unique().tolist() if "匹配方式" in grp.columns else []
            match_reasons = []
            seen_reasons: Set[str] = set()
            if "匹配方式" in grp.columns and "匹配原因" in grp.columns:
                for _, row in grp.iterrows():
                    reason = "完全相同" if str(row.get("匹配方式", "")).strip() == "完全相同" else str(row.get("匹配原因", "")).strip()
                    if reason and reason not in seen_reasons:
                        seen_reasons.add(reason)
                        match_reasons.append(reason)
            summary_row["匹配方式"] = " | ".join(match_types)
            summary_row["匹配原因"] = " | ".join(match_reasons)
    else:
        summary_row["跨表原因"] = cross_reason
    summary_row["来源文件"] = "、".join(grp["来源文件"].drop_duplicates().tolist())
    summary_row["对应表内第几条"] = build_position_summary(grp, compare_mode)
    summary_row["代表SQL"] = grp["SQL语句"].iloc[0]
    for col in enrich_columns or []:
        summary_row[col] = merge_group_values(grp[col]) if col in grp.columns else ""
    summary_row["规则判断是否有大表"] = "是" if matched_tables else "否"
    summary_row["规则判断涉及到的大表名称"] = "、".join(matched_tables)
    if include_performance_rules:
        performance_cause, performance_basis, optimization_hint = infer_performance_cause(
            sql=str(grp["SQL语句"].iloc[0]),
            sql_type=str(grp["SQL类型"].iloc[0]),
            matched_tables=matched_tables,
            large_table_map=large_table_map,
        )
        summary_row["利用规则判断慢SQL原因"] = performance_cause
        summary_row["规则判断依据"] = performance_basis
        summary_row["规则判断优化方向"] = optimization_hint
    return summary_row


def clean_col_name(value: object) -> str:
    text = str(value)
    text = (
        text.replace("\ufeff", "")
        .replace("\u200b", "")
        .replace("\u200c", "")
        .replace("\u200d", "")
        .replace("\u00a0", " ")
        .replace("\u3000", " ")
    )
    text = re.sub(r"\s+", "", text)
    return text.strip()


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
            html_errors = []
            for encoding in ("utf-8", "gb18030", "gbk", None):
                try:
                    if encoding is None:
                        tables = pd.read_html(file_path, header=None)
                    else:
                        tables = pd.read_html(file_path, header=None, encoding=encoding)
                    return tables[0]
                except Exception as exc:
                    html_errors.append(f"{encoding or 'auto'}: {exc}")
            raise ValueError("无法按 HTML 方式读取该 .xls 文件；尝试编码失败：" + " | ".join(html_errors))

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


def select_active_enrich_columns(
    configured_columns: Optional[List[str]],
    dataframes: Iterable[pd.DataFrame],
) -> List[str]:
    if not configured_columns:
        return []

    available_columns: Set[str] = set()
    for df in dataframes:
        available_columns.update(str(col) for col in df.columns)

    return [col for col in configured_columns if col in available_columns]


def build_similarity_reports(
    file1: str,
    file2: str,
    df1: pd.DataFrame,
    df2: pd.DataFrame,
    start_row1: int,
    start_row2: int,
    large_table_map: Optional[Dict[str, int]] = None,
    enrich_columns: Optional[List[str]] = None,
    ignore_whitespace: bool = True,
    deduplicate_within_file: bool = False,
    include_performance_rules: bool = False,
    compare_mode: str = COMPARE_MODE_STRICT,
    match_info_map: Optional[Dict[str, Dict[int, Dict[str, object]]]] = None,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    detail_rows: List[Dict[str, object]] = []
    file1_name = Path(file1).name
    file2_name = Path(file2).name
    files = [(file1_name, df1, start_row1), (file2_name, df2, start_row2)]

    for file_name, df, start_row in files:
        other_file_name = file2_name if file_name == file1_name else file1_name
        for idx, row in df.reset_index(drop=True).iterrows():
            raw_sql = row.get(SQL_COLUMN_NAME)
            sql = normalize_sql(raw_sql, ignore_whitespace=False)
            if not sql:
                continue
            compare_sql_key = normalize_sql(raw_sql, ignore_whitespace=ignore_whitespace)
            strict_signature = build_similarity_signature(sql, ignore_whitespace=ignore_whitespace)
            date_signature = build_date_only_signature(sql, ignore_whitespace=ignore_whitespace)
            inlist_signature = build_inlist_signature(sql, ignore_whitespace=ignore_whitespace)
            parameter_signature = build_parameter_signature(sql, ignore_whitespace=ignore_whitespace)
            strict_class_id = "S" + hashlib.md5(strict_signature.encode("utf-8")).hexdigest()[:8].upper()
            raw_fingerprint = row.get("指纹", "") if "指纹" in df.columns else ""
            detail_rows.append(
                {
                    "来源文件": file_name,
                    "对应表内第几条": int(row.get("_source_seq", idx + 1)),
                    "原始Excel行号": int(row.get("_excel_row_no", start_row + idx)),
                    "相似类ID": strict_class_id,
                    "SQL类型": classify_sql(sql),
                    "原表指纹": raw_fingerprint,
                    "SQL语句": sql,
                    "相似SQL特征": strict_signature,
                    "日期归一特征": date_signature,
                    "IN归一特征": inlist_signature,
                    "参数归一特征": parameter_signature,
                    "_compare_sql_key": compare_sql_key,
                    "SQL涉及表": "、".join(extract_table_names(sql)),
                }
            )
            if compare_mode == COMPARE_MODE_SMART and match_info_map and file_name in match_info_map:
                match_info = match_info_map[file_name].get(int(row.get("_source_seq", idx + 1)), {})
                detail_rows[-1]["匹配方式"] = match_info.get("match_type", "")
                detail_rows[-1]["匹配原因"] = match_info.get("match_reason", "")
                detail_rows[-1]["对方表第几条"] = match_info.get("other_positions", "")
            for col in enrich_columns or []:
                detail_rows[-1][col] = row.get(col, "") if col in df.columns else ""

    detail_df = pd.DataFrame(detail_rows)
    if detail_df.empty:
        return pd.DataFrame(), pd.DataFrame()

    if compare_mode == COMPARE_MODE_SMART and "对方表第几条" in detail_df.columns:
        smart_group_map = build_smart_group_ids_from_detail(detail_df)
        if smart_group_map:
            detail_df["相似类ID"] = detail_df.apply(
                lambda row: smart_group_map.get(f"{row['来源文件']}#{row['对应表内第几条']}", row["相似类ID"]),
                axis=1,
            )

    summary_source_df = detail_df.copy()
    if deduplicate_within_file:
        summary_source_df = summary_source_df.drop_duplicates(
            subset=["来源文件", "_compare_sql_key"],
            keep="first",
        ).copy()
    summary_rows = []
    for class_id, grp in summary_source_df.groupby("相似类ID", sort=False):
        if grp["来源文件"].nunique() < 2:
            summary_rows.append(
                build_summary_row(
                    grp,
                    class_id,
                    "仅单表出现",
                    "否",
                    large_table_map,
                    enrich_columns,
                    include_performance_rules,
                    compare_mode,
                )
            )
            continue

        if compare_mode == COMPARE_MODE_SMART:
            exact_overlap = set()
            file_names = list(grp["来源文件"].drop_duplicates())
            if len(file_names) >= 2:
                file_a = grp[grp["来源文件"] == file_names[0]]
                file_b = grp[grp["来源文件"] == file_names[1]]
                exact_overlap = set(file_a["SQL语句"].astype(str)) & set(file_b["SQL语句"].astype(str))

            summary_rows.append(
                build_summary_row(
                    grp,
                    class_id,
                    "",
                    "是" if exact_overlap else "否",
                    large_table_map,
                    enrich_columns,
                    include_performance_rules,
                    compare_mode,
                )
            )
            continue

        file_names = list(grp["来源文件"].drop_duplicates())
        file_a = grp[grp["来源文件"] == file_names[0]]
        file_b = grp[grp["来源文件"] == file_names[1]]
        exact_overlap = set(file_a["SQL语句"].astype(str)) & set(file_b["SQL语句"].astype(str))

        consumed_indexes: Set[int] = set()
        if exact_overlap:
            exact_grp = grp[grp["SQL语句"].astype(str).isin(exact_overlap)].copy()
            if not exact_grp.empty:
                summary_rows.append(
                    build_summary_row(
                        exact_grp,
                        class_id,
                        "跨表完全相同",
                        "是",
                        large_table_map,
                        enrich_columns,
                        include_performance_rules,
                        compare_mode,
                    )
                )
                consumed_indexes.update(set(exact_grp.index.tolist()))

        remaining_grp = grp.loc[~grp.index.isin(consumed_indexes)].copy()
        if not remaining_grp.empty:
            if exact_overlap:
                summary_rows.append(
                    build_summary_row(
                        remaining_grp,
                        class_id,
                        "单表内完全相同",
                        "否",
                        large_table_map,
                        enrich_columns,
                        include_performance_rules,
                        compare_mode,
                    )
                )
            else:
                has_exact_match, cross_reason = infer_cross_file_reason(remaining_grp)
                summary_rows.append(
                    build_summary_row(
                        remaining_grp,
                        class_id,
                        cross_reason,
                        has_exact_match,
                        large_table_map,
                        enrich_columns,
                        include_performance_rules,
                        compare_mode,
                    )
                )

    summary_df = pd.DataFrame(summary_rows).sort_values(
        ["重复条数", "涉及文件数"],
        ascending=[False, False],
    )

    detail_df = detail_df.drop(columns=["_compare_sql_key"])
    return summary_df, detail_df


def compare_sql_files(
    file1: str,
    file2: str,
    output_dir: str,
    ignore_whitespace: bool = True,
    deduplicate_within_file: bool = False,
    table_stats_file: str = "",
    large_table_threshold: int = 1000000,
    enrich_columns: Optional[List[str]] = None,
    include_performance_rules: bool = False,
    compare_mode: str = COMPARE_MODE_STRICT,
) -> Tuple[str, Dict[str, int], pd.DataFrame, pd.DataFrame]:
    os.makedirs(output_dir, exist_ok=True)

    df1 = load_excel(file1)
    df2 = load_excel(file2)

    df1, start_row1 = fix_header(df1, SQL_COLUMN_NAME)
    df2, start_row2 = fix_header(df2, SQL_COLUMN_NAME)

    df1 = df1[df1[SQL_COLUMN_NAME].notna()].copy()
    df2 = df2[df2[SQL_COLUMN_NAME].notna()].copy()
    df1["_source_seq"] = range(1, len(df1) + 1)
    df2["_source_seq"] = range(1, len(df2) + 1)
    df1["_excel_row_no"] = range(start_row1, start_row1 + len(df1))
    df2["_excel_row_no"] = range(start_row2, start_row2 + len(df2))
    raw_df1 = df1.copy()
    raw_df2 = df2.copy()
    detail_df1 = df1.copy()
    detail_df2 = df2.copy()
    active_enrich_columns = select_active_enrich_columns(enrich_columns, [detail_df1, detail_df2])

    df1["_compare_sql_key"] = df1[SQL_COLUMN_NAME].apply(lambda x: normalize_sql(x, ignore_whitespace))
    df2["_compare_sql_key"] = df2[SQL_COLUMN_NAME].apply(lambda x: normalize_sql(x, ignore_whitespace))

    if deduplicate_within_file:
        df1 = df1.drop_duplicates(subset=["_compare_sql_key"], keep="first").copy()
        df2 = df2.drop_duplicates(subset=["_compare_sql_key"], keep="first").copy()
    else:
        # Keep the helper column for non-deduplicated filtering, then drop it from exported sheets later.
        pass

    helper_columns = ["_compare_sql_key", "_source_seq", "_excel_row_no"]
    match_info_1: Dict[int, Dict[str, object]] = {}
    match_info_2: Dict[int, Dict[str, object]] = {}
    if compare_mode == COMPARE_MODE_SMART:
        match_info_1 = _build_match_info(df1, df2, ignore_whitespace=ignore_whitespace, compare_mode=compare_mode)
        match_info_2 = _build_match_info(df2, df1, ignore_whitespace=ignore_whitespace, compare_mode=compare_mode)
        only1_df = _attach_match_columns(df1, match_info_1, matched=False)
        only2_df = _attach_match_columns(df2, match_info_2, matched=False)
        both1_df = _attach_match_columns(df1, match_info_1, matched=True)
        both2_df = _attach_match_columns(df2, match_info_2, matched=True)
    else:
        keys1 = set(df1["_compare_sql_key"])
        keys2 = set(df2["_compare_sql_key"])
        only1_keys = keys1 - keys2
        only2_keys = keys2 - keys1
        both_keys = keys1 & keys2
        only1_df = df1[df1["_compare_sql_key"].isin(only1_keys)].copy()
        only2_df = df2[df2["_compare_sql_key"].isin(only2_keys)].copy()
        both1_df = df1[df1["_compare_sql_key"].isin(both_keys)].copy()
        both2_df = df2[df2["_compare_sql_key"].isin(both_keys)].copy()

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
        df1=detail_df1,
        df2=detail_df2,
        start_row1=start_row1,
        start_row2=start_row2,
        large_table_map=large_table_map,
        enrich_columns=active_enrich_columns,
        ignore_whitespace=ignore_whitespace,
        deduplicate_within_file=deduplicate_within_file,
        include_performance_rules=include_performance_rules,
        compare_mode=compare_mode,
        match_info_map={
            Path(file1).name: match_info_1,
            Path(file2).name: match_info_2,
        } if compare_mode == COMPARE_MODE_SMART else None,
    )

    if compare_mode == COMPARE_MODE_SMART and "匹配方式" in similarity_summary_df.columns:
        only1_df, both1_df = _move_single_file_summary_rows_from_both_to_only(
            only1_df,
            both1_df,
            Path(file1).name,
            similarity_summary_df,
            similarity_detail_df,
        )
        only2_df, both2_df = _move_single_file_summary_rows_from_both_to_only(
            only2_df,
            both2_df,
            Path(file2).name,
            similarity_summary_df,
            similarity_detail_df,
        )
        single_file_reason_map: Dict[Tuple[str, int], str] = {}
        single_file_group_map: Dict[Tuple[str, int], str] = {}
        single_summary = similarity_summary_df[
            similarity_summary_df["匹配方式"].astype(str) == "单表内匹配"
        ].copy()
        if not single_summary.empty:
            reason_by_group = {
                str(row["相似类ID"]): str(row["匹配原因"])
                for _, row in single_summary.iterrows()
            }
            detail_subset = similarity_detail_df[
                similarity_detail_df["相似类ID"].astype(str).isin(reason_by_group.keys())
            ]
            members_by_group: Dict[str, List[Tuple[str, int]]] = {}
            for _, row in detail_subset.iterrows():
                group_id = str(row["相似类ID"])
                key = (str(row["来源文件"]), int(row["对应表内第几条"]))
                single_file_reason_map[key] = reason_by_group[group_id]
                members_by_group.setdefault(group_id, []).append(key)
            for group_id, members in members_by_group.items():
                grouped_by_file: Dict[str, List[int]] = {}
                for file_name, seq in members:
                    grouped_by_file.setdefault(file_name, []).append(seq)
                text_parts = []
                for file_name, seqs in grouped_by_file.items():
                    seq_text = "、".join(str(x) for x in sorted(seqs))
                    text_parts.append(f"{file_name}: {seq_text}")
                summary_text = " | ".join(text_parts)
                for key in members:
                    single_file_group_map[key] = summary_text
        only1_df = _apply_single_file_match_labels(only1_df, Path(file1).name, single_file_reason_map, single_file_group_map)
        only2_df = _apply_single_file_match_labels(only2_df, Path(file2).name, single_file_reason_map, single_file_group_map)
        only1_df = _collapse_single_file_match_rows(only1_df)
        only2_df = _collapse_single_file_match_rows(only2_df)
        both1_df = _apply_group_summary_to_shared_sheet(
            both1_df,
            Path(file1).name,
            similarity_summary_df,
            similarity_detail_df,
        )
        both2_df = _apply_group_summary_to_shared_sheet(
            both2_df,
            Path(file2).name,
            similarity_summary_df,
            similarity_detail_df,
        )
        both1_df = _collapse_shared_sheet_rows(
            both1_df,
            Path(file1).name,
            similarity_summary_df,
            similarity_detail_df,
        )
        both2_df = _collapse_shared_sheet_rows(
            both2_df,
            Path(file2).name,
            similarity_summary_df,
            similarity_detail_df,
        )
        only1_df = only1_df.drop(columns=helper_columns)
        only2_df = only2_df.drop(columns=helper_columns)
        only1_df = _drop_only_sheet_helper_columns(only1_df)
        only2_df = _drop_only_sheet_helper_columns(only2_df)
        both1_df = _apply_group_summary_to_shared_sheet(
            both1_df,
            Path(file1).name,
            similarity_summary_df,
            similarity_detail_df,
        )
        both2_df = _apply_group_summary_to_shared_sheet(
            both2_df,
            Path(file2).name,
            similarity_summary_df,
            similarity_detail_df,
        )
        both1_df = _finalize_shared_sheet_columns(both1_df)
        both2_df = _finalize_shared_sheet_columns(both2_df)
        both1_df = both1_df.drop(columns=helper_columns)
        both2_df = both2_df.drop(columns=helper_columns)
    elif compare_mode != COMPARE_MODE_SMART:
        only1_df, both1_df = _move_strict_single_side_rows_from_both_to_only(
            only1_df,
            both1_df,
            Path(file1).name,
            similarity_summary_df,
            similarity_detail_df,
        )
        only2_df, both2_df = _move_strict_single_side_rows_from_both_to_only(
            only2_df,
            both2_df,
            Path(file2).name,
            similarity_summary_df,
            similarity_detail_df,
        )
        only1_df, both1_df = _sync_shared_summary_rows(
            only1_df,
            both1_df,
            Path(file1).name,
            similarity_summary_df,
            similarity_detail_df,
        )
        only2_df, both2_df = _sync_shared_summary_rows(
            only2_df,
            both2_df,
            Path(file2).name,
            similarity_summary_df,
            similarity_detail_df,
        )
        single_file_reason_map, single_file_group_map = _build_single_file_maps_from_summary(
            similarity_summary_df,
            similarity_detail_df,
        )
        only1_df = _apply_single_file_match_labels(
            only1_df,
            Path(file1).name,
            single_file_reason_map,
            single_file_group_map,
        )
        only2_df = _apply_single_file_match_labels(
            only2_df,
            Path(file2).name,
            single_file_reason_map,
            single_file_group_map,
        )
        only1_df = _collapse_single_file_match_rows(only1_df)
        only2_df = _collapse_single_file_match_rows(only2_df)
        only1_df = only1_df.drop(columns=helper_columns)
        only2_df = only2_df.drop(columns=helper_columns)
        only1_df = _drop_only_sheet_helper_columns(only1_df)
        only2_df = _drop_only_sheet_helper_columns(only2_df)
        both1_df = _apply_group_summary_to_shared_sheet(
            both1_df,
            Path(file1).name,
            similarity_summary_df,
            similarity_detail_df,
        )
        both2_df = _apply_group_summary_to_shared_sheet(
            both2_df,
            Path(file2).name,
            similarity_summary_df,
            similarity_detail_df,
        )
        both1_df = _collapse_shared_sheet_rows(
            both1_df,
            Path(file1).name,
            similarity_summary_df,
            similarity_detail_df,
        )
        both2_df = _collapse_shared_sheet_rows(
            both2_df,
            Path(file2).name,
            similarity_summary_df,
            similarity_detail_df,
        )
        both1_df = _finalize_shared_sheet_columns(both1_df)
        both2_df = _finalize_shared_sheet_columns(both2_df)
        both1_df = both1_df.drop(columns=helper_columns)
        both2_df = both2_df.drop(columns=helper_columns)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        raw_df1.drop(columns=["_source_seq", "_excel_row_no"]).to_excel(writer, sheet_name=f"原表_{name1}"[:31], index=False)
        raw_df2.drop(columns=["_source_seq", "_excel_row_no"]).to_excel(writer, sheet_name=f"原表_{name2}"[:31], index=False)
        only1_df.to_excel(writer, sheet_name=sheet_only1, index=False)
        only2_df.to_excel(writer, sheet_name=sheet_only2, index=False)
        both1_df.to_excel(writer, sheet_name=sheet_both1, index=False)
        both2_df.to_excel(writer, sheet_name=sheet_both2, index=False)
        _hide_internal_export_columns(similarity_summary_df).to_excel(writer, sheet_name="相似SQL归类汇总", index=False)
        _hide_internal_export_columns(similarity_detail_df).to_excel(writer, sheet_name="相似SQL明细", index=False)

    stats = {
        "表1总数": len(df1),
        "表2总数": len(df2),
        "仅表1": len(only1_df),
        "仅表2": len(only2_df),
        "共有": len(both1_df),
        "相似类数": len(similarity_summary_df),
        "大表阈值": large_table_threshold,
        "表内去重": "是" if deduplicate_within_file else "否",
        "对比方式": "智能匹配" if compare_mode == COMPARE_MODE_SMART else "严格匹配",
    }
    return output_path, stats, similarity_summary_df, similarity_detail_df


class SqlDiffApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("SQL 语句 Excel 比较工具")
        self.root.geometry("680x700")
        self.root.minsize(680, 500)

        self.file1_var = tk.StringVar()
        self.file2_var = tk.StringVar()
        self.table_stats_var = tk.StringVar()
        self.use_table_stats_var = tk.BooleanVar(value=False)
        self.output_dir_var = tk.StringVar(value=str(Path.home() / "Desktop"))
        self.large_table_threshold_var = tk.StringVar(value="1000000")
        self.relaxed_match_var = tk.BooleanVar(value=True)
        self.include_performance_rules_var = tk.BooleanVar(value=False)
        self.compare_mode_var = tk.StringVar(value=COMPARE_MODE_STRICT)
        self.status_var = tk.StringVar(value="请选择两个 Excel 文件和输出目录。")
        self.status_text: Optional[tk.Text] = None
        self.history_enrich_columns = load_enrich_columns_config()

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
        self._build_table_stats_row(frame)
        self._build_file_row(frame, "输出目录", self.output_dir_var, self.choose_output_dir, select_file=False)

        option_frame = ttk.Frame(frame)
        option_frame.pack(fill="x", pady=(18, 10))

        ttk.Label(
            option_frame,
            text="对比选项",
            font=("Microsoft YaHei UI", 10, "bold"),
        ).pack(anchor="w")

        compare_mode_row = ttk.Frame(option_frame)
        compare_mode_row.pack(fill="x", pady=(0, 6))
        ttk.Radiobutton(
            compare_mode_row,
            text="严格匹配",
            variable=self.compare_mode_var,
            value=COMPARE_MODE_STRICT,
        ).pack(side="left")
        ttk.Radiobutton(
            compare_mode_row,
            text="智能匹配",
            variable=self.compare_mode_var,
            value=COMPARE_MODE_SMART,
        ).pack(side="left", padx=(12, 0))

        ttk.Checkbutton(
            option_frame,
            text="合并判断相同SQL（忽略空白差异，并合并同一文件内重复SQL）",
            variable=self.relaxed_match_var,
        ).pack(anchor="w")

        ttk.Checkbutton(
            option_frame,
            text="利用规则判断慢SQL原因（在相似SQL归类汇总中增加原因/依据/优化方向三列）",
            variable=self.include_performance_rules_var,
        ).pack(anchor="w", pady=(6, 0))

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
        ttk.Button(button_frame, text="汇总字段设置", command=self.open_enrich_columns_dialog).pack(side="left", padx=(12, 0))

        status_title = ttk.Label(frame, text="运行状态", font=("Microsoft YaHei UI", 10, "bold"))
        status_title.pack(anchor="w", pady=(8, 6))

        status_frame = ttk.Frame(frame)
        status_frame.pack(fill="both", expand=True)

        self.status_text = tk.Text(
            status_frame,
            wrap="word",
            height=12,
            relief="solid",
            bd=1,
            padx=10,
            pady=10,
        )
        scrollbar = ttk.Scrollbar(status_frame, orient="vertical", command=self.status_text.yview)
        self.status_text.configure(yscrollcommand=scrollbar.set)
        self.status_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self._set_status("请选择两个 Excel 文件和输出目录。")

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

    def _build_table_stats_row(self, parent: ttk.Frame) -> None:
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=8)

        ttk.Label(row, text="表数据量统计", width=12).pack(side="left")
        ttk.Checkbutton(
            row,
            text="是否上传",
            variable=self.use_table_stats_var,
            command=self._toggle_table_stats_state,
        ).pack(side="left", padx=(0, 10))
        self.table_stats_entry = ttk.Entry(row, textvariable=self.table_stats_var, state="disabled")
        self.table_stats_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.table_stats_button = ttk.Button(
            row,
            text="选择文件",
            command=self.choose_table_stats_file,
            state="disabled",
        )
        self.table_stats_button.pack(side="left")

    def _toggle_table_stats_state(self) -> None:
        enabled = self.use_table_stats_var.get()
        state = "normal" if enabled else "disabled"
        self.table_stats_entry.configure(state=state)
        self.table_stats_button.configure(state=state)
        if not enabled:
            self.table_stats_var.set("")

    def _set_status(self, message: str) -> None:
        self.status_var.set(message)
        if self.status_text is None:
            return
        self.status_text.configure(state="normal")
        self.status_text.delete("1.0", "end")
        self.status_text.insert("1.0", message)
        self.status_text.see("end")
        self.status_text.configure(state="disabled")

    def _append_status(self, message: str) -> None:
        if self.status_text is None:
            self.status_var.set(message)
            return
        self.status_text.configure(state="normal")
        current = self.status_text.get("1.0", "end-1c").strip()
        if current:
            self.status_text.insert("end", f"\n\n{message}")
        else:
            self.status_text.insert("1.0", message)
        self.status_text.see("end")
        self.status_text.configure(state="disabled")
        self.status_var.set(message)

    def open_enrich_columns_dialog(self) -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title("汇总字段设置")
        dialog.geometry("520x420")
        dialog.minsize(420, 320)
        dialog.transient(self.root)
        dialog.grab_set()

        container = ttk.Frame(dialog, padding=16)
        container.pack(fill="both", expand=True)

        ttk.Label(
            container,
            text="维护“相似SQL归类汇总”需要额外带出的表头名",
            font=("Microsoft YaHei UI", 11, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            container,
            text="这些字段如果在上传文件里存在，就会带到汇总表；不存在则自动留空，不会报错。",
            foreground="#666666",
            justify="left",
        ).pack(anchor="w", pady=(6, 12))

        list_frame = ttk.Frame(container)
        list_frame.pack(fill="both", expand=True)

        listbox = tk.Listbox(list_frame)
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=listbox.yview)
        scrollbar.pack(side="right", fill="y")
        listbox.configure(yscrollcommand=scrollbar.set)

        for col in self.history_enrich_columns:
            listbox.insert("end", col)

        entry_row = ttk.Frame(container)
        entry_row.pack(fill="x", pady=(12, 8))
        new_col_var = tk.StringVar()
        ttk.Entry(entry_row, textvariable=new_col_var).pack(side="left", fill="x", expand=True)

        def add_column() -> None:
            col = new_col_var.get().strip()
            if not col:
                return
            existing = list(listbox.get(0, "end"))
            if col in existing:
                messagebox.showwarning("重复字段", f"字段“{col}”已经存在。", parent=dialog)
                return
            listbox.insert("end", col)
            new_col_var.set("")

        ttk.Button(entry_row, text="新增字段", command=add_column).pack(side="left", padx=(10, 0))

        action_row = ttk.Frame(container)
        action_row.pack(fill="x", pady=(0, 8))

        def remove_selected() -> None:
            selection = listbox.curselection()
            if not selection:
                return
            for index in reversed(selection):
                listbox.delete(index)

        def move_up() -> None:
            selection = listbox.curselection()
            if not selection or selection[0] == 0:
                return
            for index in selection:
                text = listbox.get(index)
                listbox.delete(index)
                listbox.insert(index - 1, text)
                listbox.selection_set(index - 1)

        def move_down() -> None:
            selection = listbox.curselection()
            if not selection or selection[-1] == listbox.size() - 1:
                return
            for index in reversed(selection):
                text = listbox.get(index)
                listbox.delete(index)
                listbox.insert(index + 1, text)
                listbox.selection_set(index + 1)

        def reset_default() -> None:
            listbox.delete(0, "end")
            for col in DEFAULT_HISTORY_ENRICH_COLUMNS:
                listbox.insert("end", col)

        ttk.Button(action_row, text="删除选中", command=remove_selected).pack(side="left")
        ttk.Button(action_row, text="上移", command=move_up).pack(side="left", padx=(10, 0))
        ttk.Button(action_row, text="下移", command=move_down).pack(side="left", padx=(10, 0))
        ttk.Button(action_row, text="恢复默认", command=reset_default).pack(side="left", padx=(10, 0))

        footer_row = ttk.Frame(container)
        footer_row.pack(fill="x", pady=(10, 0))

        def save_and_close() -> None:
            columns = [str(item).strip() for item in listbox.get(0, "end") if str(item).strip()]
            deduped: List[str] = []
            seen: Set[str] = set()
            for col in columns:
                if col not in seen:
                    seen.add(col)
                    deduped.append(col)
            self.history_enrich_columns = deduped
            save_enrich_columns_config(self.history_enrich_columns)
            self._append_status(f"已更新汇总附加字段配置：共 {len(self.history_enrich_columns)} 个。")
            dialog.destroy()

        ttk.Button(footer_row, text="取消", command=dialog.destroy).pack(side="right")
        ttk.Button(footer_row, text="保存", command=save_and_close).pack(side="right", padx=(0, 10))

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
        if not self.use_table_stats_var.get():
            self.use_table_stats_var.set(True)
            self._toggle_table_stats_state()
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
        table_stats_file = self.table_stats_var.get().strip() if self.use_table_stats_var.get() else ""
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

        if self.use_table_stats_var.get() and not table_stats_file:
            messagebox.showwarning("缺少文件", "你已选择上传表数据量统计，请再选择对应文件。")
            return

        self._set_status("正在比较，请稍候...")
        self._append_status(f"表1文件：{file1}")
        self._append_status(f"表2文件：{file2}")
        self._append_status(f"是否上传表数据量统计：{'是' if self.use_table_stats_var.get() else '否'}")
        if table_stats_file:
            self._append_status(f"表数据量统计文件：{table_stats_file}")
            self._append_status(f"大表阈值：{large_table_threshold}")
        self._append_status(f"对比方式：{'智能匹配' if self.compare_mode_var.get() == COMPARE_MODE_SMART else '严格匹配'}")
        self._append_status(f"合并判断相同SQL：{'是' if self.relaxed_match_var.get() else '否'}")
        self._append_status(f"利用规则判断慢SQL原因：{'是' if self.include_performance_rules_var.get() else '否'}")
        self.root.update_idletasks()

        try:
            relaxed_match = self.relaxed_match_var.get()
            output_name = f"SQL比较结果_{Path(file1).stem}_VS_{Path(file2).stem}.xlsx"
            output_path = os.path.join(output_dir, output_name)
            if os.path.exists(output_path):
                self._append_status(f"检测到同名结果文件：{output_path}")
                should_overwrite = messagebox.askyesno(
                    "文件已存在",
                    f"结果文件已存在：\n{output_path}\n\n是否覆盖？",
                )
                if not should_overwrite:
                    self._set_status("已取消生成：存在同名结果文件，且未选择覆盖。")
                    return

            output_path, stats, _, _ = compare_sql_files(
                file1=file1,
                file2=file2,
                output_dir=output_dir,
                ignore_whitespace=relaxed_match,
                deduplicate_within_file=relaxed_match,
                table_stats_file=table_stats_file,
                large_table_threshold=large_table_threshold,
                enrich_columns=self.history_enrich_columns,
                include_performance_rules=self.include_performance_rules_var.get(),
                compare_mode=self.compare_mode_var.get(),
            )
            self._append_status("比较处理完成，正在整理结果信息...")
        except BadZipFile:
            msg = (
                "文件扩展名看起来像 Excel，但文件内容不是标准的 .xlsx/.xlsm 格式。\n"
                "请确认不是把 CSV、截图导出文件或临时文件误当成 Excel 选进来了。"
            )
            self._set_status(f"比较失败：{msg}")
            messagebox.showerror("比较失败", msg)
            return
        except Exception as exc:
            self._set_status(f"比较失败：{exc}")
            traceback.print_exc()
            messagebox.showerror("比较失败", f"{exc}")
            return

        stats_file_desc = table_stats_file if table_stats_file else "未上传"
        self._set_status(
            "比较完成。\n"
            f"结果文件：{output_path}\n\n"
            f"表1总数：{stats['表1总数']}\n"
            f"表2总数：{stats['表2总数']}\n"
            f"仅表1：{stats['仅表1']}\n"
            f"仅表2：{stats['仅表2']}\n"
            f"共有：{stats['共有']}\n"
            f"相似SQL类数：{stats['相似类数']}\n"
            f"对比方式：{stats['对比方式']}\n"
            f"合并判断相同SQL：{'是' if relaxed_match else '否'}\n"
            f"表内去重：{stats['表内去重']}\n"
            f"是否上传表数据量统计：{'是' if self.use_table_stats_var.get() else '否'}\n"
            f"表数据量统计文件：{stats_file_desc}\n"
            f"大表阈值：{stats['大表阈值']}\n\n"
            f"利用规则判断慢SQL原因：{'是' if self.include_performance_rules_var.get() else '否'}\n"
            f"汇总附加字段数：{len(self.history_enrich_columns)}\n\n"
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
            "4. “严格匹配”只把完全匹配的 SQL 视为共有；“智能匹配”会额外识别仅参数不同、仅日期不同、IN列表不同、分页条件不同、SELECT字段增加/减少、WHERE条件增加/减少，以及高相似度匹配。\n"
            "5. “合并判断相同SQL”勾选后，会同时忽略空白差异，并合并同一文件内重复 SQL；不勾选时，严格按原始 SQL 比较，并保留表内重复行。\n"
            "6. 如果上传“表数据量统计”，程序会根据大表阈值判断相似 SQL 是否涉及大表。\n"
            "7. 可以通过“汇总字段设置”按钮，自定义相似SQL归类汇总要额外带出的表头名。\n\n"
            "二、相似 SQL 归类规则\n"
            "1. 相似归类不是按业务语义，而是按标准化后的 SQL 骨架分组。\n"
            "2. 归类时会统一大小写和空白格式。\n"
            "3. 字符串常量会替换成 ?str?。\n"
            "4. 日期会替换成 ?date?，8 位日期串会替换成 ?date8?。\n"
            "5. 数字会替换成 ?num?。\n"
            "6. IN (...) 会折叠成 in(?list?)，VALUES (...) 会折叠成 values(?vals?)。\n"
            "7. 导出的结果表不会显示内部用的分组编号，直接通过“来源文件 / 对应表内第几条 / 对方表第几条 / 组内匹配条目”等字段来说明分组关系。\n\n"
            "三、汇总表字段说明\n"
            "1. 是否存在完全相同SQL：表示两个表里是否出现过完全一致的原始 SQL。\n"
            "2. 严格匹配模式下，汇总表主要看“跨表原因”；智能匹配模式下，汇总表主要看“匹配方式 / 匹配原因”。\n"
            "3. 对应表内第几条：显示这一类 SQL 分别出现在各原表里的第几条，方便回原表定位。\n"
            "4. 共有(表1) / 共有(表2) 会只保留当前表这一组里的第一条，并在“对方表第几条”里展示整组对应关系。\n"
            "5. 仅表1 / 仅表2 里如果存在同表内多条归成一组，会显示“匹配方式 / 匹配原因 / 组内匹配条目”。\n"
            "6. 如果勾选“利用规则判断慢SQL原因”，会额外生成“利用规则判断慢SQL原因 / 规则判断依据 / 规则判断优化方向”三列；这三列基于 SQL 结构特征做规则推断，用于辅助归类，不代表数据库已实锤慢因。\n"
            "7. 规则判断是否有大表 / 规则判断涉及到的大表名称：基于“表数据量统计”文件中的表名和记录数判断。"
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
