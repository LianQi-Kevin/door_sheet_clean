#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Door digits checker & summary (single-file, config-first).

保持现有功能/口径不变：
- 文本归一化：NFKC + 空白折叠 + 常见占位串视为缺失（避免 "nan"/"<NA>" 注入）；
- 忽略备注关键字（字面量匹配，NA 不命中）；
- 饰面同义词映射、token 化去重、连接符；
- “宽口径”判定：备注/饰面在分组内若唯一 → 只输出标题，不列“门编号”；多值 → 按值分块并列门清单；
- 行级 digits 校验：expected-first 生成 + 严格边界匹配 + 期望型号替换/追加；
- 输出：地上/地下楼层（若有）+ digits行级校验。

调用方式（程序化）：
    from pathlib import Path
    from door_digits_single import run
    run(Path("input.xlsx"), Path("output.xlsx"), only_errors=False)
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Dict, Tuple, Optional

import pandas as pd
import unicodedata
from pandas._libs.missing import NAType

# ========================== 配置（文件头、全大写） ==========================
# 文本归一化
NORMALIZE_FORM: str = "NFKC"  # 可选: "NFC"/"NFKC"/"NFD"/"NFKD"
CASEFOLD_TEXT: bool = False  # 如需不区分大小写匹配，可设 True

# 常见占位串（归一化后以缺失处理）
MISSING_STRINGS: set[str] = {"", "nan", "none", "null", "<na>", "n/a", "na"}

# 备注忽略（字面量包含；NA 不命中）
IGNORE_REMARK_SUBSTRINGS: List[str] = ["门槛"]
IGNORE_REMARK_CASE_SENSITIVE: bool = True  # True=区分大小写

# 饰面同义词映射 & token 连接符
FINISH_SYNONYMS: Dict[str, str] = {"喷涂": "静电粉末喷涂", "深灰色静电粉末喷涂": "静电粉末喷涂",
                                   "WD-21木饰面": "木纹转印", "WD-24木饰面": "木纹转印", "WD-01木饰面": "木纹转印",
                                   "木纹": "木纹转印", }
FINISH_TOKEN_JOINER: str = " "  # 例如可改为 "、"

# 备注唯一/多值判定口径（当前为“宽口径”，即只要有效值唯一即可不列门编号）
STRICT_REMARK_UNIQUENESS: bool = False  # 预留开关：True=严格口径（全组一致才省略门编号）


# 楼层分组配置
@dataclass(frozen=True)
class LevelGroup:
    name: str
    levels: List[str]
    rename_map: Dict[str, str]
    ordered: List[str]
    sheet_name: str

    def filter_df(self, df: pd.DataFrame) -> pd.DataFrame:
        return df[df["层号"].isin(self.levels)]

    def ordered_columns(self, cols: Iterable[str]) -> List[str]:
        return [c for c in self.ordered if c in cols]


BASEMENT_GROUP = LevelGroup(name="basement", levels=["B1", "B2", "B3"],
                            rename_map={"B1": "地下一层门数量", "B2": "地下二层门数量", "B3": "地下三层门数量"},
                            ordered=["类型", "型号", "洞口尺寸(宽X高)", "开启方式", "地下一层门数量", "地下二层门数量",
                                     "地下三层门数量",
                                     "合计", "备注"], sheet_name="地下楼层", )

ABOVE_GROUP = LevelGroup(name="above",
                         levels=["F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "2#机房层"],
                         rename_map={"F1": "一层门数量", "F2": "二层门数量", "F3": "三层门数量", "F4": "四层门数量",
                                     "F5": "五层门数量",
                                     "F6": "六层门数量", "F7": "七层门数量", "F8": "八层门数量", "F9": "九层门数量",
                                     "F10": "十层门数量",
                                     "F11": "十一层门数量", "2#机房层": "机房层门数量"},
                         ordered=["类型", "型号", "洞口尺寸(宽X高)", "开启方式", "一层门数量", "二层门数量",
                                  "三层门数量", "四层门数量",
                                  "五层门数量", "六层门数量", "七层门数量", "八层门数量", "九层门数量", "十层门数量",
                                  "十一层门数量",
                                  "机房层门数量", "合计", "备注"], sheet_name="地上楼层", )

LEVEL_GROUPS: List[LevelGroup] = [BASEMENT_GROUP, ABOVE_GROUP]


# ========================== 文本/通用工具 ==========================
def normalize_text(value: object, nf: str = NORMALIZE_FORM, casefold: bool = CASEFOLD_TEXT) -> str | NAType:
    """
    归一化文本：NFKC + 空白折叠 + 常见占位串视为缺失；必要时大小写规约。
    返回：规范化后的 str；若视为缺失则返回 pd.NA。
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return pd.NA
    s = str(value).strip()
    if not s:
        return pd.NA
    s = unicodedata.normalize(nf, s)
    s = s.replace("（", "(").replace("）", ")")
    s = re.sub(r"\s+", " ", s).strip()
    if s.casefold() in MISSING_STRINGS:
        return pd.NA
    if casefold:
        s = s.casefold()
    return s if s else pd.NA


def fmt_cell(x: object) -> Optional[str]:
    """门编号/层号安全格式化；缺失返回 None（避免 'nan' 注入拼装文本）。"""
    nx = normalize_text(x)
    if nx is pd.NA:
        return None
    return str(nx)


def filter_remarks_for_ignore(series: pd.Series) -> pd.Series:
    """
    忽略备注关键字（字面量包含、可选大小写；NA 不命中）。
    返回：归一化+忽略后的 Series。
    """
    s = series.map(lambda x: normalize_text(x))
    if not IGNORE_REMARK_SUBSTRINGS:
        return s
    mask = pd.Series(False, index=s.index)
    for kw in IGNORE_REMARK_SUBSTRINGS:
        if not kw:
            continue
        mask |= s.str.contains(kw, case=IGNORE_REMARK_CASE_SENSITIVE, regex=False, na=False)
    return s.mask(mask, other=pd.NA)


def split_finish_tokens(x: object, synonym_map: Dict[str, str] | None = None) -> List[str]:
    """饰面分词：按常见分隔符拆分 → 同义词映射 → 去重（保序）。"""
    s = normalize_text(x)
    if s is pd.NA:
        return []
    parts = re.split(r"[、/;；,，\s]+", s)
    toks, seen = [], set()
    for p in parts:
        p = p.strip()
        if not p:
            continue
        if synonym_map:
            p = synonym_map.get(p, p)
        if p not in seen:
            seen.add(p)
            toks.append(p)
    return toks


def stable_join_tokens(tokens: List[str], joiner: str = FINISH_TOKEN_JOINER) -> str:
    """稳定连接 tokens（不排序，保出现顺序）。"""
    return joiner.join(tokens)


def parse_mm_series(s: pd.Series) -> pd.Series:
    """宽/高清洗：数字化→四舍五入→可空 Int64。"""
    return pd.to_numeric(s, errors="coerce").round().astype("Int64")


# ========================== 备注/饰面块拼装 ==========================
def _build_block_unique_or_grouped(values_series: pd.Series, title: str, group_rows: pd.DataFrame,
                                   basement_levels: Iterable[str]) -> str:
    """
    通用块构造：唯一→只标题；多值→标题+门清单。
    values_series：每行已标准化的“展示值”（备注或饰面），允许 NA。
    """
    valid = values_series.dropna()
    if valid.empty:
        return ""

    uniq_values = pd.Series(valid).unique()  # 保序、无排序
    # 宽口径：只要有效值唯一即可不列门编号；严格口径（可选）
    if (not STRICT_REMARK_UNIQUENESS and len(uniq_values) == 1) or (
            STRICT_REMARK_UNIQUENESS and len(valid.unique()) == 1 and valid.size == len(group_rows)):
        return f"{title}: {uniq_values[0]}"

    basement_set = set(str(x).strip() for x in basement_levels)
    blocks: list[str] = []
    # 将标准化值回写后分组
    df2 = group_rows.copy()
    colname = f"__{title}_norm__"
    df2[colname] = values_series
    for v, grp in df2.dropna(subset=[colname]).groupby(colname, sort=False):
        vtxt = fmt_cell(v)
        if not vtxt:
            continue
        # 组内门编号清单
        doors_list: list[str] = []
        for _, r in grp.iterrows():
            num = fmt_cell(r.get("门编号"))
            level = fmt_cell(r.get("层号"))
            if num is None:
                continue
            if level in basement_set and level is not None:
                doors_list.append(f"{num}（{level}层）")
            else:
                doors_list.append(num)
        doors_list = pd.unique(doors_list).tolist()
        if not doors_list:
            continue
        blocks.append(f"{title}: {vtxt}\n门编号: " + "、".join(doors_list))
    return "\n\n".join(blocks)


def build_finish_note(group: pd.DataFrame, basement_levels: Iterable[str]) -> str:
    """饰面块：token 化→同义词→去重→展示值；唯一→只标题， 多值→标题+门清单。"""
    if "饰面" not in group.columns:
        return ""
    tok_series = group["饰面"].apply(lambda v: split_finish_tokens(v, synonym_map=FINISH_SYNONYMS))
    disp_series = tok_series.apply(lambda toks: stable_join_tokens(toks, FINISH_TOKEN_JOINER) if toks else pd.NA)
    return _build_block_unique_or_grouped(disp_series, "饰面", group, basement_levels)


def build_note_original_like(group: pd.DataFrame, basement_levels: Iterable[str]) -> str:
    """备注块：归一化→忽略关键字→唯一/多值分支（与饰面一致口径）。"""
    if "备注" not in group.columns:
        return ""
    remarks_norm = filter_remarks_for_ignore(group["备注"])
    return _build_block_unique_or_grouped(remarks_norm, "备注", group, basement_levels)


def build_note_with_finish(group: pd.DataFrame, basement_levels: Iterable[str]) -> str:
    """合并备注块与饰面块（以空行分隔），任一为空则只返回另一个。"""
    note_block = build_note_original_like(group, basement_levels)
    finish_block = build_finish_note(group, basement_levels)
    if note_block and finish_block:
        return f"{note_block}\n\n{finish_block}"
    return note_block or finish_block or ""


# ========================== 数据清洗与汇总 ==========================
def clean_base_keep_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    基础清洗：列名去空格；门型→型号；关键文本列使用 StringDtype 并做基本规范化；
    保留原始列（不做无关列删除）。
    """
    df = df.copy()
    df.columns = df.columns.str.replace(" ", "", regex=False)

    if "门型" in df.columns:
        df["门型"] = df["门型"].astype("string")
        df["门型"] = (
            df["门型"].str.replace("（", "(", regex=False).str.replace(
                "）", ")", regex=False).str.replace("’", "'", regex=False).str.replace(
                "‘", "'", regex=False))
    df = df.rename(columns={"门样式": "类型", "门型": "型号"})

    # 统一文本列为 StringDtype，保持 <NA> 语义；避免 astype(str) 产生 'nan'
    for c in [c for c in ("类型", "型号", "层号", "门编号", "备注", "门扇", "饰面") if c in df.columns]:
        df[c] = df[c].astype("string")
        df[c] = df[c].str.replace(" ", "", regex=False).str.strip()

    return df


def summarize_group(df_clean: pd.DataFrame, group: LevelGroup, basement_levels_for_note: Iterable[str]) -> pd.DataFrame:
    """
    单组（地上/地下）汇总：洞口尺寸、开启方式、楼层计数、备注拼装与列顺序。
    """
    df_sub = group.filter_df(df_clean).copy()
    if df_sub.empty:
        return pd.DataFrame(columns=group.ordered)

    # 洞口尺寸 & 开启方式
    w = parse_mm_series(df_sub["宽度"])
    h = parse_mm_series(df_sub["高度"])
    df_sub["洞口尺寸(宽X高)"] = pd.Series(
        [f"{int(wi)}X{int(hi)}" if pd.notna(wi) and pd.notna(hi) else pd.NA for wi, hi in zip(w, h)], dtype="string")
    df_sub["开启方式"] = df_sub["门扇"].map({"单扇": "单扇平开", "双扇": "双扇平开"}).astype("string")

    # 计数透视
    counts = (df_sub.groupby(["类型", "型号", "层号"]).size().unstack(fill_value=0).reset_index().sort_values(
        ["类型", "型号"]))

    # 合计及重命名
    level_cols_present = [c for c in group.levels if c in counts.columns]
    counts["合计"] = counts[level_cols_present].sum(axis=1) if level_cols_present else 0
    counts = counts.rename(columns=group.rename_map)
    level_cols_renamed = [group.rename_map.get(c, c) for c in level_cols_present]
    if level_cols_renamed:
        counts[level_cols_renamed] = counts[level_cols_renamed].replace(0, "")

    # 备注/饰面汇总
    extras = (df_sub.groupby(["类型", "型号"], sort=False).apply(lambda g: pd.Series(
        {"洞口尺寸(宽X高)": "，".join(pd.Series(g["洞口尺寸(宽X高)"].dropna().astype("string")).unique()),
         "开启方式": "，".join(pd.Series(g["开启方式"].dropna().astype("string")).unique()),
         "备注": build_note_with_finish(g, basement_levels_for_note), })).reset_index())

    result = counts.merge(extras, on=["类型", "型号"], how="left")
    return result.reindex(columns=group.ordered_columns(result.columns))


# ========================== 行级 digits 校验 ==========================
PAT_ANY_DIGITS = re.compile(r"(?<!\d)D[0-9'\u2019]+(?!\d)")


def build_expected_digits(width_mm: object, height_mm: object, pad_to: int = 2) -> Tuple[str, str]:
    """
    由真实宽/高(mm)生成期望 digits：
      expected_with_D: "Dxx'yy'"
      expected_core  : "xx'yy'"
    宽/高 < 1000mm 时各自补零到 pad_to 位；非整百追加英文单引号 '。
    """
    if pd.isna(width_mm) or pd.isna(height_mm):
        return "", ""
    w = int(width_mm)
    h = int(height_mm)
    w_code = str(w // 100).zfill(pad_to)
    h_code = str(h // 100).zfill(pad_to)
    w_q = "'" if (w % 100) else ""
    h_q = "'" if (h % 100) else ""
    core = f"{w_code}{w_q}{h_code}{h_q}"
    return f"D{core}", core


def make_expected_model(model: str, expected_with_D: str) -> str:
    """在型号中将旧 digits 段整体替换为期望 digits；没有则尾部追加。"""
    model = str(model or "")
    if not expected_with_D:
        return model
    if PAT_ANY_DIGITS.search(model):
        return PAT_ANY_DIGITS.sub(expected_with_D, model)
    return f"{model} {expected_with_D}".strip()


def build_row_level_expected_first(df: pd.DataFrame) -> pd.DataFrame:
    """
    行级 digits 校验表（match + reason_code/reason_detail + 期望型号）：
    - match=True  : 型号包含严格边界匹配的期望 digits
    - match=False : 能生成期望 digits，但型号里没找到
    - match=pd.NA : 宽/高缺失或无法生成期望 digits
    """
    work = df.copy()
    work["宽度_mm"] = parse_mm_series(work["宽度"])
    work["高度_mm"] = parse_mm_series(work["高度"])

    exp = work.apply(lambda r: pd.Series(build_expected_digits(r["宽度_mm"], r["高度_mm"], pad_to=2),
                                         index=["expected_with_D", "expected_core"]), axis=1)
    work = pd.concat([work, exp], axis=1)

    def _did_match(row: pd.Series) -> bool:
        if not row["expected_with_D"]:
            return False
        pattern = rf"(?<!\d){re.escape(row['expected_with_D'])}(?!\d)"
        return bool(re.search(pattern, str(row["型号"])))

    def classify(row: pd.Series):
        if pd.isna(row["宽度_mm"]) or pd.isna(row["高度_mm"]):
            return pd.NA, "wh_missing", "宽度/高度为空或无法解析"
        if not row["expected_with_D"]:
            return pd.NA, "expected_build_fail", "未能按宽高生成期望 digits"
        if _did_match(row):
            return True, "ok", ""
        return False, "expected_not_found", "型号未包含以真实尺寸派生的期望 digits，请手动核对/修正"

    work[["match", "reason_code", "reason_detail"]] = work.apply(classify, axis=1, result_type="expand")
    work["期望型号"] = work.apply(lambda r: make_expected_model(r["型号"], r["expected_with_D"]), axis=1)

    cols = [c for c in ["层号" if "层号" in work.columns else None, "门编号" if "门编号" in work.columns else None,
                        "类型" if "类型" in work.columns else None, "型号", "期望型号", "宽度_mm", "高度_mm",
                        "expected_with_D",
                        "expected_core", "match", "reason_code", "reason_detail",
                        "备注" if "备注" in work.columns else None] if
            c is not None]
    return work[cols]


# ========================== 顶层入口（程序化调用） ==========================
def run(input_path: Path, output_path: Path, only_errors: bool = False) -> None:
    """
    读取 Excel → 清洗 → 汇总（地上/地下）→ digits行级校验 → 写出。
    参数:
        input_path  : 输入 Excel 路径
        output_path : 输出 Excel 路径
        only_errors : True 时，digits 行级校验仅输出 match != True 的行
    """
    df_raw = pd.read_excel(input_path, sheet_name=0, header=0)
    df_clean = clean_base_keep_columns(df_raw)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # 楼层组汇总
        for grp in LEVEL_GROUPS:
            tbl = summarize_group(df_clean, grp, basement_levels_for_note=BASEMENT_GROUP.levels)
            if not tbl.empty:
                tbl.to_excel(writer, sheet_name=grp.sheet_name, index=False)

        # digits 行级校验
        row_level = build_row_level_expected_first(df_clean)
        if only_errors:
            row_level = row_level[row_level["match"].ne(True)]
        row_level.to_excel(writer, sheet_name="digits行级校验", index=False)


if __name__ == '__main__':
    input_file = Path(r"sample_excel.xlsx")
    output_file = Path(r"target_excel.xlsx")
    run(input_file, output_file, only_errors=False)
    print(f"Processed {input_file} and saved results to {output_file}.")
