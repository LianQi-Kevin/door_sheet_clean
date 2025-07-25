#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Door digits checker (expected-first, unified LevelGroup)

功能概要
--------
1) 期望优先（expected-first）：
   - 由真实 宽/高(mm) 生成期望 digits：D{w百毫米[+'']}{h百毫米[+'']}
     * 宽/高 < 1000mm 时各自 zfill(2) 补零
     * 非整百在对应段后加英文单引号 '
   - 在型号中用严格边界 (?<!\\d) ... (?!\\d) 搜索该期望 digits
   - 命中则 match=True，否则 False；若宽/高缺失或无法生成期望 digits，则 match=pd.NA
   - 同时保留 reason_code / reason_detail 说明不匹配原因

2) 地上/地下汇总通过 LevelGroup 统一配置与逻辑，减少重复代码。

3) CLI：
   - 未传 --sheet → 读取首个 sheet（sheet_name=0）
   - 未传 --log-level → 默认 INFO
   - --only-errors 只输出 match != True 的行（行级 sheet）

依赖
----
- pandas
- openpyxl
"""

from __future__ import annotations

import argparse
import logging
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional, Tuple, Dict, List

import pandas as pd

# ----------------------- Defaults -----------------------
DEFAULT_INPUT: Path = Path("sample_excel.xlsx")
DEFAULT_OUTPUT: Path = Path("target_excel.xlsx")
DEFAULT_SHEET: Optional[str] = None  # None → 读首个 sheet
DEFAULT_LOG_LEVEL: str = "INFO"

# 统一替换型号中旧 digits 的正则（把任何 D[0-9'…] 替换为新的期望 digits）
PAT_ANY_DIGITS = re.compile(r"(?<!\d)D[0-9'\u2019]+(?!\d)")

logger = logging.getLogger(__name__)


# ----------------------- LevelGroup 配置 -----------------------
@dataclass
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
                                     "F11": "十一层门数量", "2#机房层": "机房层门数量", },
                         ordered=["类型", "型号", "开启方式", "洞口尺寸(宽X高)", "一层门数量", "二层门数量",
                                  "三层门数量", "四层门数量",
                                  "五层门数量", "六层门数量", "七层门数量", "八层门数量", "九层门数量", "十层门数量",
                                  "十一层门数量",
                                  "机房层门数量", "合计", "备注"], sheet_name="地上楼层", )

LEVEL_GROUPS: List[LevelGroup] = [BASEMENT_GROUP, ABOVE_GROUP]


# ----------------------- Logging -----------------------
def configure_logging(level: str = DEFAULT_LOG_LEVEL) -> None:
    """Quick root logger config (官方文档推荐用法)."""
    logging.basicConfig(level=getattr(logging, level.upper(), logging.INFO),
                        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s", )
    logger.debug("logging configured at %s", level)


# ----------------------- Common utils -----------------------
def _is_missing(val: object) -> bool:
    if pd.isna(val):
        return True
    s = str(val).strip()
    return s in ("", "nan", "NaN")


def build_note(group: pd.DataFrame, basement_levels: Iterable[str]) -> str:
    """聚合备注（地下层门编号追加“（B?层）”），保持你原有逻辑。"""
    basement_set = set(basement_levels)
    valid = group[~group["备注"].apply(_is_missing)]
    if valid.empty:
        return ""
    if valid["备注"].nunique() == 1 and len(valid) == len(group):
        return str(valid["备注"].iloc[0]).strip()

    blocks: list[str] = []
    for remark, grp in valid.groupby("备注", sort=False):
        lines = [f"备注: {str(remark).strip()}"]
        doors = (grp.apply(lambda r: (
            f"{str(r['门编号']).strip()}（{r['层号']}层）" if str(r["层号"]).strip() in basement_set else str(
                r["门编号"]).strip()), axis=1, ).unique())
        lines.append(f"门编号: {'、'.join(doors)}")
        blocks.append("\n".join(lines))
    return "\n\n".join(blocks)


def clean_base(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.str.replace(" ", "", regex=False)
    for col in df.columns:
        df[col] = df[col].astype(str).str.replace(" ", "", regex=False)
    df = df.drop(columns=["防火门区域", "房间功能", "防火级别", "区域"], errors="ignore")
    if "门型" in df.columns:
        df["门型"] = (
            df["门型"].str.replace("（", "(", regex=False).str.replace(
                "）", ")", regex=False).str.replace("’", "'", regex=False).str.replace(
                "‘", "'", regex=False))
    return df.rename(columns={"门样式": "类型", "门型": "型号"})


# ----------------------- expected-first 主逻辑 -----------------------
def build_expected_digits(width_mm: object, height_mm: object, pad_to: int = 2) -> Tuple[str, str]:
    """
    由真实宽/高(mm)生成期望 digits：
      expected_with_D: "Dxx'yy'"
      expected_core  : "xx'yy'"
    宽/高 < 1000mm 时各自补零到 2 位；非整百追加英文单引号 '。
    """
    if pd.isna(width_mm) or pd.isna(height_mm):
        return "", ""
    w, h = int(width_mm), int(height_mm)
    w_code = str(w // 100).zfill(pad_to)
    h_code = str(h // 100).zfill(pad_to)
    w_q = "'" if (w % 100) else ""
    h_q = "'" if (h % 100) else ""
    core = f"{w_code}{w_q}{h_code}{h_q}"
    return f"D{core}", core


def make_expected_model(model: str, expected_with_D: str) -> str:
    """将型号中旧 digits 段整体替换为期望 digits；没有则直接追加。"""
    model = str(model or "")
    if PAT_ANY_DIGITS.search(model):
        return PAT_ANY_DIGITS.sub(expected_with_D, model)
    return f"{model} {expected_with_D}".strip()


def build_row_level_expected_first(df: pd.DataFrame, basement_levels: Iterable[str]) -> pd.DataFrame:
    """
    生成 digits 行级校验表（保留 match + reason_code/reason_detail）：
    - match=True  : 型号中找到了严格边界匹配的期望 digits
    - match=False : 能生成期望 digits，但型号里没找到它
    - match=pd.NA : 宽/高缺失或无法生成期望 digits
    """
    work = df.copy()

    # 1) 宽/高清洗
    work["宽度_mm"] = pd.to_numeric(work["宽度"], errors="coerce").round().astype("Int64")
    work["高度_mm"] = pd.to_numeric(work["高度"], errors="coerce").round().astype("Int64")

    # 2) 由真实宽高中构造期望 digits
    exp = work.apply(lambda r: pd.Series(build_expected_digits(r["宽度_mm"], r["高度_mm"], pad_to=2),
                                         index=["expected_with_D", "expected_core"]), axis=1)
    work = pd.concat([work, exp], axis=1)

    # 3) 分类 & match 判定（内部使用严格边界正则，但不输出）
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

    # 4) 期望型号（建议应改成什么）
    work["期望型号"] = work.apply(lambda r: make_expected_model(r["型号"], r["expected_with_D"]), axis=1)

    # 5) 日志（match=True→DEBUG，其余→WARNING）
    for idx, r in work.iterrows():
        msg = (f"row={idx} 型号={r.get('型号')} 宽高=({r.get('宽度_mm')}x{r.get('高度_mm')}) "
               f"expected={r.get('expected_with_D')} match={r.get('match')} reason={r.get('reason_code')}")
        (logger.debug if r["match"] is True else logger.warning)(msg)

    # 6) 输出列
    remark_col = "备注" if "备注" in work.columns else None
    cols = [c for c in ["层号" if "层号" in work.columns else None, "门编号" if "门编号" in work.columns else None,
                        "类型" if "类型" in work.columns else None, "型号", "期望型号", "宽度_mm", "高度_mm",
                        "expected_with_D", "expected_core", "match", "reason_code", "reason_detail", remark_col,
                        ] if c is not None]
    return work[cols]


# ----------------------- 统一楼层组汇总 -----------------------
def summarize_group(df_clean: pd.DataFrame, group: LevelGroup, basement_levels_for_note: Iterable[str]) -> pd.DataFrame:
    df_sub = group.filter_df(df_clean).copy()
    if df_sub.empty:
        return pd.DataFrame(columns=group.ordered)

    # 洞口尺寸 & 开启方式
    w = pd.to_numeric(df_sub["宽度"], errors="coerce").round().astype("Int64")
    h = pd.to_numeric(df_sub["高度"], errors="coerce").round().astype("Int64")
    df_sub["洞口尺寸(宽X高)"] = w.astype(str) + "X" + h.astype(str)
    df_sub["开启方式"] = df_sub["门扇"].map({"单扇": "单扇平开", "双扇": "双扇平开"})

    # 计数透视
    counts = (df_sub.groupby(["类型", "型号", "层号"]).size().unstack(fill_value=0).reset_index().sort_values(
        ["类型", "型号"]))

    # 合计（只统计本组包含且真实存在的列）
    level_cols_present = [c for c in group.levels if c in counts.columns]
    counts["合计"] = counts[level_cols_present].sum(axis=1) if level_cols_present else 0

    # 重命名并替换 0 为 ""（仅楼层列）
    counts = counts.rename(columns=group.rename_map)
    level_cols_renamed = [group.rename_map.get(c, c) for c in level_cols_present]
    if level_cols_renamed:
        counts[level_cols_renamed] = counts[level_cols_renamed].replace(0, "")

    # 附加尺寸、开启方式、备注
    extras = (df_sub.groupby(["类型", "型号"], sort=False).apply(lambda g: pd.Series(
        {"洞口尺寸(宽X高)": "，".join(pd.Series(g["洞口尺寸(宽X高)"].dropna().astype(str)).unique()),
         "开启方式": "，".join(pd.Series(g["开启方式"].dropna().astype(str)).unique()),
         "备注": build_note(g, basement_levels_for_note), })).reset_index())

    result = counts.merge(extras, on=["类型", "型号"], how="left")
    return result.reindex(columns=group.ordered_columns(result.columns))


# ----------------------- CLI -----------------------
def parse_args(argv: Optional[Iterable[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Door digits checker (expected-first, unified groups).",
                                     formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument("-i", "--input", type=Path, default=DEFAULT_INPUT, help="Input Excel path.")
    parser.add_argument("-o", "--output", type=Path, default=DEFAULT_OUTPUT, help="Output Excel path.")
    parser.add_argument("-s", "--sheet", type=str, default=DEFAULT_SHEET,
                        help="Sheet name to read. Omit to read the first sheet.")
    parser.add_argument("--log-level", type=str, default=DEFAULT_LOG_LEVEL,
                        help="Logging level (DEBUG, INFO, WARNING, ERROR).")
    parser.add_argument("--only-errors", action="store_true",
                        help="Only keep rows where match != True in row-level sheet.")
    return parser.parse_args(argv)


def main(argv: Optional[Iterable[str]] = None) -> None:
    args = parse_args(argv)
    configure_logging(args.log_level)

    # --sheet 未给 → 读首个 sheet（pandas: sheet_name=0）
    sheet_to_read: int | str = 0 if args.sheet is None else args.sheet
    logger.info("reading %s (sheet=%s)", args.input, sheet_to_read)
    df_raw = pd.read_excel(args.input, sheet_name=sheet_to_read, header=0)

    df_clean = clean_base(df_raw)

    # 行级校验
    row_level = build_row_level_expected_first(df_clean, basement_levels=BASEMENT_GROUP.levels)
    if args.only_errors:
        row_level = row_level[row_level["match"].ne(True)]

    # 分组汇总
    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        for group in LEVEL_GROUPS:
            tbl = summarize_group(df_clean, group, basement_levels_for_note=BASEMENT_GROUP.levels)
            if not tbl.empty:
                tbl.to_excel(writer, sheet_name=group.sheet_name, index=False)
        row_level.to_excel(writer, sheet_name="digits行级校验", index=False)

    logger.info("done -> %s", args.output)


if __name__ == "__main__":
    main()
