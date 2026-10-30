# -*- coding: utf-8 -*-
"""
EME 实验数据管理系统 - Excel 留痕生成（v3：仅数据 + 真实公式，不内嵌图表）

工作表结构：
  总表    元数据 + 各物质汇总（回收率/D→M/M→A 的均值与RSD，真实公式）
  RT表   各样本各物质保留时间原始数据
  电压表  时间/电流两列（空表头，用户后续粘贴）
  M{组}  每个样品组一表：物质为行，A-1/A-2/A-3 与 S-1/S-2/S-3 峰面积 + 均值/RSD（真实公式）

设计原则：
- 所有"计算"列都写成真实 Excel 公式（引用同表单元格），用户改值即可重算。
- 元数据写在固定位置，方便程序回读。
"""
from __future__ import annotations
from pathlib import Path
from typing import Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

try:
    from . import config
    from .parser import ParseResult, parse_sample_name
    from .calculator import (
        aggregate_qc, aggregate_sample_phase, compute_metrics, metrics_to_dataframe,
        compute_phase_distribution
    )
except ImportError:
    import config
    from parser import ParseResult, parse_sample_name
    from calculator import (
        aggregate_qc, aggregate_sample_phase, compute_metrics, metrics_to_dataframe,
        compute_phase_distribution
    )


# ---------- 样式 ----------
HEADER_FILL = PatternFill("solid", fgColor="4472C4")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
META_KEY_FONT = Font(bold=True, color="333333")
WARN_FILL = PatternFill("solid", fgColor="FFC7CE")     # 红（超阈/超限）
WARN_FONT = Font(color="9C0006")
THIN = Side(border_style="thin", color="CCCCCC")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _style_header(ws, row: int, ncols: int):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = BORDER


def _col(n: int) -> str:
    return get_column_letter(n)


# ---------- 总表 ----------

def write_summary_sheet(wb: Workbook, res: ParseResult, metrics, meta: dict,
                        Va: float, Vd: float, use_sample: bool,
                        gcond: Optional[pd.DataFrame] = None, default_qc: Optional[dict] = None):
    """写总表：元数据区 + 组条件表（原"各物质汇总"位置；数据矩阵见「数据汇总」表）"""
    ws = wb.create_sheet("总表", 0)

    # ===== 元数据区（A1 起）=====
    ws["A1"] = "EME 实验数据总表"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells("A1:F1")

    meta_rows = [
        ("实验日期", meta.get("exp_date", res.file_date or "")),
        ("温度(℃)", meta.get("temperature", "")),
        ("湿度(%)", meta.get("humidity", "")),
        ("液相平衡柱压", meta.get("column_pressure", "")),
        ("受体相体积Va(μL)", Va),
        ("供体相体积Vd(μL)", Vd),
        ("受体相标准品名", res.qc_names.get("std_acceptor", "")),
        ("供体相标准品名", res.qc_names.get("std_donor", "")),
        ("受体相空白名", res.qc_names.get("blk_acceptor", "")),
        ("供体相空白名", res.qc_names.get("blk_donor", "")),
        ("是否使用样品(供体相)", "是" if use_sample else "否"),
        ("实验说明", meta.get("purpose", "")),
    ]
    r = 3
    for k, v in meta_rows:
        ws.cell(row=r, column=1, value=k).font = META_KEY_FONT
        ws.cell(row=r, column=2, value=v)
        r += 1

    # ===== 组条件表（原"各物质汇总"位置）=====
    cond_start = r + 2
    ws.cell(row=cond_start, column=1,
            value="各组萃取条件（数据矩阵见「数据汇总」表；明细见各 M 分表）").font = Font(bold=True, size=12)
    n_cond_rows = 0
    if gcond is not None and len(gcond):
        n_cond_rows = int((gcond["grp"] != 0).sum())
        _write_group_conditions_rows(ws, cond_start + 1, gcond)

    # ===== 标准品峰面积汇总（供体相/受体相标准品各进样针的均值+RSD；RSD>15% 标红）=====
    std_start = cond_start + 2 + n_cond_rows + 2  # 标题行（组条件表头1行 + 组行数 + 空2行）
    ws.cell(row=std_start, column=1,
            value="标准品峰面积汇总（供体相/受体相标准品，剔除针已扣除）").font = Font(bold=True, size=12)
    _std_headers = ["物质", "供体相标准品峰面积平均值", "RSD（%）",
                    "受体相标准品峰面积平均值", "RSD（%）"]
    _hr = std_start + 1
    for _i, _h in enumerate(_std_headers, start=1):
        ws.cell(row=_hr, column=_i, value=_h)
    _style_header(ws, _hr, len(_std_headers))
    _rr = _hr + 1
    for _a in res.analytes:
        _dm, _drsd, _dvals = aggregate_qc(res.df, "std_donor", _a, sample_name=(default_qc or {}).get("std_donor"))
        _am, _arsd, _avals = aggregate_qc(res.df, "std_acceptor", _a, sample_name=(default_qc or {}).get("std_acceptor"))
        ws.cell(row=_rr, column=1, value=_a).border = BORDER
        # 显式列号（2/3=供体相均值与RSD，4/5=受体相）：enumerate 步进是 1，会错位覆盖
        for _ci, _mean, _rsd, _vals in [(2, _dm, _drsd, _dvals), (4, _am, _arsd, _avals)]:
            if _vals:
                ws.cell(row=_rr, column=_ci, value=round(_mean, 3))
                _r_cell = ws.cell(row=_rr, column=_ci + 1, value=round(_rsd, 2))
                if _rsd > 15.0:
                    _r_cell.fill = WARN_FILL
                    _r_cell.font = WARN_FONT
            for _cc in (_ci, _ci + 1):
                ws.cell(row=_rr, column=_cc).border = BORDER
                ws.cell(row=_rr, column=_cc).alignment = CENTER
        _rr += 1
    for _ci, _w in [(1, 22), (2, 24), (3, 12), (4, 24), (5, 12)]:
        ws.column_dimensions[_col(_ci)].width = max(ws.column_dimensions[_col(_ci)].width or 0, _w)


def _write_group_conditions_rows(ws, header_row: int, gcond: pd.DataFrame):
    """在指定工作表指定行写"组条件"表（表头+各组一行），供总表复用。"""
    import json as _json

    def _j(col):
        v = gcond.iloc[0].get(col) if len(gcond) else None  # 仅占位，实际逐行取
        return None

    headers = ["组", "组别名", "主SLM", "SLM添加剂", "供体相pH", "供体相成分", "供体相缓冲剂浓度",
               "供体相添加剂", "受体相pH", "受体相成分", "受体相缓冲剂浓度",
               "受体相添加剂", "萃取时间(分钟)", "萃取电压(V)",
               "孔位", "异常标签", "异常孔位", "备注"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=i, value=h)
    _style_header(ws, header_row, len(headers))

    def _parse(col, row_vals):
        s = row_vals.get(col)
        if s is None:
            return ""
        s = str(s)
        if s.startswith("["):
            try:
                return ", ".join(str(x) for x in _json.loads(s))
            except Exception:
                return s
        return s

    rr = header_row + 1
    for _, r in gcond.iterrows():
        if r["grp"] == 0:
            continue
        vals = [f"M{int(r['grp'])}", r.get("group_label") or "", r.get("slm_main") or "",
                _parse("slm_additives", r), r.get("pH_donor") or "", r.get("donor_composition") or "",
                r.get("donor_buffer_conc") or "", _parse("donor_additives", r),
                r.get("pH_acceptor") or "",
                r.get("acceptor_composition") or "", r.get("acceptor_buffer_conc") or "",
                _parse("acceptor_additives", r),
                r.get("extract_time") or "", r.get("extract_voltage") or "",
                _parse("well_position", r), _parse("note_tags", r),
                _parse("abnormal_wells", r), r.get("group_note") or ""]
        has_abn = bool(str(vals[15]).strip())
        for i, v in enumerate(vals, start=1):
            cell = ws.cell(row=rr, column=i, value=v)
            cell.border = BORDER
            cell.alignment = CENTER
            if has_abn:
                cell.fill = WARN_FILL
        rr += 1
    widths = [6, 18, 14, 14, 10, 12, 14, 14, 10, 12, 14, 14, 12, 12, 16, 12, 14, 24]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[_col(i)].width = max(ws.column_dimensions[_col(i)].width or 0, w)


# ---------- M{组} 分表（含真实公式）----------

def write_group_sheet(wb: Workbook, res: ParseResult, group: int,
                      Va: float, Vd: float, use_sample: bool, excluded_samples: Optional[set] = None,
                      gcond: Optional[pd.DataFrame] = None, qc_cfg: Optional[dict] = None):
    """写单个样品组的分表：物质为行，A/S 三重复峰面积 + 均值/RSD（真实公式）。
    被剔除的原始单元格标灰；末尾追加"修正后"均值/RSD/回收率/转移率列；
    最右侧追加本组的孔位/异常标签/异常孔位/备注（来自组级条件）。
    """
    excluded_samples = excluded_samples or set()
    ws = wb.create_sheet(f"M{group}")
    # (组, 相, 重复) → 真实样品名（样品名格式可自定义，不拼接）
    _name_of = {}
    _smp_all = res.df[res.df["category"] == "sample"]
    for _n, _g_, _ph_, _r_ in zip(_smp_all["sample"], _smp_all["group"], _smp_all["phase"], _smp_all["replicate"]):
        try:
            _name_of[(int(_g_), str(_ph_), int(_r_))] = _n
        except Exception:
            continue
    GRAY_FILL = PatternFill("solid", fgColor="D9D9D9")  # 被剔除标灰
    # 本组条件（孔位/标签/备注）
    grp_cond = None
    if gcond is not None and len(gcond):
        sub = gcond[gcond["grp"] == group]
        if len(sub):
            grp_cond = sub.iloc[0]

    def _jval(col):
        if grp_cond is None:
            return ""
        v = grp_cond.get(col)
        if v is None:
            return ""
        s = str(v)
        if s.startswith("["):
            try:
                import json as _json
                return ", ".join(str(x) for x in _json.loads(s))
            except Exception:
                return s
        return s

    # 表头
    n_rep = res.n_replicates
    # 列布局：
    # A | B..   物质
    # A-1 A-2 A-3 | A均值 | A_RSD% || S-1 S-2 S-3 | S均值 | S_RSD% || 受体空白 | 供体标准品 | 供体空白 | 校正系数 | 回收率% | D→M% | M→A%
    headers = ["物质"]
    for i in range(1, n_rep + 1):
        headers.append(f"受体-{i}")
    headers += ["受体峰面积平均值", "受体峰面积RSD%"]
    for i in range(1, n_rep + 1):
        headers.append(f"供体-{i}")
    headers += ["供体峰面积平均值", "供体峰面积RSD%"]
    headers += ["受体空白平均值", "供体标准品平均值", "供体空白平均值", "受体标准品平均值",
                "校正系数", "回收率%", "供体相至SLM转移率%", "SLM至受体相转移率%"]
    # 修正列（剔除异常后）
    headers += ["修正-受体平均值", "修正-受体RSD%", "修正-供体平均值", "修正-供体RSD%",
                "修正-回收率%", "修正-供体相至SLM转移率%", "修正-SLM至受体相转移率%"]
    # 本组条件列（孔位/异常标签/异常孔位/备注，随数据行重复填值，便于筛选留档）
    headers += ["孔位", "异常标签", "异常孔位", "备注"]

    for i, h in enumerate(headers, start=1):
        ws.cell(row=1, column=i, value=h)
    _style_header(ws, 1, len(headers))

    # 找各列索引
    def col_of(name: str) -> int:
        return headers.index(name) + 1

    a_rep_cols = [col_of(f"受体-{i}") for i in range(1, n_rep + 1)]
    a_mean_col = col_of("受体峰面积平均值")
    a_rsd_col = col_of("受体峰面积RSD%")
    s_rep_cols = [col_of(f"供体-{i}") for i in range(1, n_rep + 1)]
    s_mean_col = col_of("供体峰面积平均值")
    s_rsd_col = col_of("供体峰面积RSD%")
    blk_a_col = col_of("受体空白平均值")
    std_d_col = col_of("供体标准品平均值")
    blk_d_col = col_of("供体空白平均值")
    std_a_col = col_of("受体标准品平均值")
    corr_col = col_of("校正系数")
    rec_col = col_of("回收率%")
    dm_col = col_of("供体相至SLM转移率%")
    ma_col = col_of("SLM至受体相转移率%")

    rr = 2
    for ai, analyte in enumerate(res.analytes):
        ws.cell(row=rr, column=1, value=analyte).border = BORDER

        # —— A 相三重复峰面积（原始数据；被剔除的标灰）——
        a_vals = []
        a_kept = []  # 未剔除的值（用于修正计算）
        for rep in range(1, n_rep + 1):
            sample_name = _name_of.get((group, "A", rep))
            sub = res.df[(res.df["sample"] == sample_name) & (res.df["analyte"] == analyte) & (res.df["category"] == "sample")] if sample_name else res.df.iloc[0:0]
            if not len(sub):
                continue  # 该组缺这个重复位（补做/漏针）→ 留空：AVERAGE/STDEV 忽略空单元格，与计算器口径一致（写 0 会拉低均值）
            v = float(sub["area"].iloc[0])
            a_vals.append(v)
            cell = ws.cell(row=rr, column=a_rep_cols[rep - 1], value=v)
            if sample_name in excluded_samples:
                cell.fill = GRAY_FILL  # 标灰表示已剔除
            else:
                a_kept.append(v)
        # A均值公式 =AVERAGE(B..)
        first = _col(a_rep_cols[0]) + str(rr)
        last = _col(a_rep_cols[-1]) + str(rr)
        ws.cell(row=rr, column=a_mean_col, value=f"=AVERAGE({first}:{last})")
        # A_RSD% 公式 =STDEV(B..)/AVERAGE(B..)*100
        ws.cell(row=rr, column=a_rsd_col,
                value=f"=IF(COUNT({first}:{last})<2,0,STDEV({first}:{last})/AVERAGE({first}:{last})*100)")

        # —— S 相三重复 ——
        s_vals = []
        s_kept = []
        for rep in range(1, n_rep + 1):
            sample_name = _name_of.get((group, "S", rep))
            sub = res.df[(res.df["sample"] == sample_name) & (res.df["analyte"] == analyte) & (res.df["category"] == "sample")]
            if not len(sub):
                continue  # 同上：缺失重复位留空
            v = float(sub["area"].iloc[0])
            s_vals.append(v)
            cell = ws.cell(row=rr, column=s_rep_cols[rep - 1], value=v)
            if sample_name in excluded_samples:
                cell.fill = GRAY_FILL
            else:
                s_kept.append(v)
        first_s = _col(s_rep_cols[0]) + str(rr)
        last_s = _col(s_rep_cols[-1]) + str(rr)
        ws.cell(row=rr, column=s_mean_col, value=f"=AVERAGE({first_s}:{last_s})")
        ws.cell(row=rr, column=s_rsd_col,
                value=f"=IF(COUNT({first_s}:{last_s})<2,0,STDEV({first_s}:{last_s})/AVERAGE({first_s}:{last_s})*100)")

        # —— 质控参考（多标准品时按该组组合过滤）——
        _qc = qc_cfg or {}
        blk_a_mean, _, _ = aggregate_qc(res.df, "blk_acceptor", analyte, sample_name=_qc.get("blk_acceptor"))
        std_d_mean, _, _ = aggregate_qc(res.df, "std_donor", analyte, sample_name=_qc.get("std_donor"))
        blk_d_mean, _, _ = aggregate_qc(res.df, "blk_donor", analyte, sample_name=_qc.get("blk_donor"))
        std_a_mean, _, _ = aggregate_qc(res.df, "std_acceptor", analyte, sample_name=_qc.get("std_acceptor"))
        ws.cell(row=rr, column=blk_a_col, value=blk_a_mean)
        ws.cell(row=rr, column=std_d_col, value=std_d_mean)
        ws.cell(row=rr, column=blk_d_col, value=blk_d_mean)
        ws.cell(row=rr, column=std_a_col, value=std_a_mean)

        # 校正系数 = 供体标准品/受体标准品
        corr_formula = f"=IF({_col(std_a_col)}{rr}=0,1,{_col(std_d_col)}{rr}/{_col(std_a_col)}{rr})"
        ws.cell(row=rr, column=corr_col, value=corr_formula)

        # 回收率公式：((A均值-受体空白)*Va)/((供体标准品-供体空白)*Vd)*校正系数*100
        a_mean_ref = f"{_col(a_mean_col)}{rr}"
        blk_a_ref = f"{_col(blk_a_col)}{rr}"
        std_d_ref = f"{_col(std_d_col)}{rr}"
        blk_d_ref = f"{_col(blk_d_col)}{rr}"
        corr_ref = f"{_col(corr_col)}{rr}"
        rec_formula = f"=IFERROR(({a_mean_ref}-{blk_a_ref})*{Va}/(({std_d_ref}-{blk_d_ref})*{Vd})*{corr_ref}*100,\"\")"
        ws.cell(row=rr, column=rec_col, value=rec_formula)

        # D→M = 100 - (S均值-供体空白)/(供体标准品-供体空白)*100
        s_mean_ref = f"{_col(s_mean_col)}{rr}"
        dm_formula = f"=IFERROR(100-({s_mean_ref}-{blk_d_ref})/({std_d_ref}-{blk_d_ref})*100,\"\")"
        if use_sample:
            ws.cell(row=rr, column=dm_col, value=dm_formula)
        else:
            ws.cell(row=rr, column=dm_col, value="未启用供体相")

        # M→A = 100 * 回收率 / D→M
        rec_ref = f"{_col(rec_col)}{rr}"
        dm_ref = f"{_col(dm_col)}{rr}"
        ma_formula = f"=IFERROR(100*{rec_ref}/{dm_ref},\"\")"
        if use_sample:
            ws.cell(row=rr, column=ma_col, value=ma_formula)
        else:
            ws.cell(row=rr, column=ma_col, value="未启用供体相")

        # —— 修正列（剔除异常后的均值/RSD/回收率/转移率）——
        import numpy as np
        corr_val = (std_d_mean / std_a_mean) if std_a_mean else 1.0
        a_mean_fix = float(np.mean(a_kept)) if a_kept else None
        a_rsd_fix = (float(np.std(a_kept, ddof=1)) / a_mean_fix * 100) if (len(a_kept) >= 2 and a_mean_fix) else None
        s_mean_fix = float(np.mean(s_kept)) if s_kept else None
        s_rsd_fix = (float(np.std(s_kept, ddof=1)) / s_mean_fix * 100) if (len(s_kept) >= 2 and s_mean_fix) else None
        # 修正回收率
        rec_fix = None
        if a_mean_fix is not None:
            denom_fix = (std_d_mean - blk_d_mean) * Vd
            if denom_fix != 0:
                rec_fix = (a_mean_fix - blk_a_mean) * Va / denom_fix * corr_val * 100.0
        # 修正 D→M / M→A
        dm_fix = None; ma_fix = None
        if use_sample and s_mean_fix is not None and (std_d_mean - blk_d_mean) != 0:
            dm_fix = 100.0 - (s_mean_fix - blk_d_mean) / (std_d_mean - blk_d_mean) * 100.0
            if dm_fix != 0 and rec_fix is not None:
                ma_fix = 100.0 * rec_fix / dm_fix
        fix_vals = [a_mean_fix, a_rsd_fix, s_mean_fix, s_rsd_fix, rec_fix, dm_fix, ma_fix]
        fix_cols = [col_of("修正-受体平均值"), col_of("修正-受体RSD%"),
                    col_of("修正-供体平均值"), col_of("修正-供体RSD%"),
                    col_of("修正-回收率%"), col_of("修正-供体相至SLM转移率%"),
                    col_of("修正-SLM至受体相转移率%")]
        for fc, fv in zip(fix_cols, fix_vals):
            ws.cell(row=rr, column=fc, value=round(fv, 2) if fv is not None else "none")

        # 本组条件列（孔位/异常标签/异常孔位/备注）
        cond_vals = [_jval("well_position"), _jval("note_tags"),
                     _jval("abnormal_wells"), _jval("group_note")]
        cond_cols = [col_of("孔位"), col_of("异常标签"), col_of("异常孔位"), col_of("备注")]
        for cc, cv in zip(cond_cols, cond_vals):
            ws.cell(row=rr, column=cc, value=cv)

        # 边框 + RSD超阈标红
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=rr, column=c)
            cell.border = BORDER
            cell.alignment = CENTER
        # A_RSD / S_RSD > 15 标红（数值由公式算出，这里无法预判，但可加条件格式）
        rr += 1

    # 加条件格式：RSD列 > 15 标红；D→M/M→A 列超出 0-100 标红
    from openpyxl.formatting.rule import CellIsRule
    last_data_row = rr - 1
    # A_RSD
    ws.conditional_formatting.add(
        f"{_col(a_rsd_col)}2:{_col(a_rsd_col)}{last_data_row}",
        CellIsRule(operator="greaterThan", formula=["15"], fill=WARN_FILL, font=WARN_FONT))
    ws.conditional_formatting.add(
        f"{_col(s_rsd_col)}2:{_col(s_rsd_col)}{last_data_row}",
        CellIsRule(operator="greaterThan", formula=["15"], fill=WARN_FILL, font=WARN_FONT))
    # D→M / M→A 超出0-100。这两列可能是文本（"未启用供体相"/IFERROR 空串），
    # 数值型 CellIsRule 对"文本>任意数"恒真会整列误标红 → 用 ISNUMBER 守卫的公式规则
    from openpyxl.formatting.rule import FormulaRule
    for cc in [dm_col, ma_col]:
        _rng = f"{_col(cc)}2:{_col(cc)}{last_data_row}"
        _rel = f"{_col(cc)}2"
        ws.conditional_formatting.add(_rng, FormulaRule(
            formula=[f"AND(ISNUMBER({_rel}),{_rel}<0)"], fill=WARN_FILL, font=WARN_FONT))
        ws.conditional_formatting.add(_rng, FormulaRule(
            formula=[f"AND(ISNUMBER({_rel}),{_rel}>100)"], fill=WARN_FILL, font=WARN_FONT))

    # 列宽
    ws.column_dimensions["A"].width = 22
    for c in range(2, len(headers) + 1):
        ws.column_dimensions[_col(c)].width = 13
    ws.freeze_panes = "B2"


# ---------- RT 表 ----------

def write_rt_sheet(wb: Workbook, res: ParseResult):
    """RT表：各样本各物质的保留时间（原始数据，供出图）"""
    ws = wb.create_sheet("RT表")
    sub = res.df[res.df["category"] == "sample"][["sample", "group", "phase", "replicate", "analyte", "rt"]].copy()
    # 透视：行=样本，列=物质
    if len(sub) == 0:
        ws["A1"] = "无样本RT数据"
        return
    pivot = sub.pivot_table(index="sample", columns="analyte", values="rt", aggfunc="first")
    # 排序样本
    _fmt_sk = getattr(res, "sample_format", "") or ""
    def _sort_key(s):
        g, ph, rp = parse_sample_name(s, pattern=_fmt_sk or None) or (99, "Z", 99)
        return (g, ph, rp)
    pivot = pivot.reindex(sorted(pivot.index, key=_sort_key))
    # 写入
    ws.cell(row=1, column=1, value="样本")
    for j, a in enumerate(pivot.columns, start=2):
        ws.cell(row=1, column=j, value=a)
    _style_header(ws, 1, len(pivot.columns) + 1)
    for i, (sample, row) in enumerate(pivot.iterrows(), start=2):
        ws.cell(row=i, column=1, value=sample).border = BORDER
        for j, a in enumerate(pivot.columns, start=2):
            v = row[a]
            cell = ws.cell(row=i, column=j, value=float(v) if pd.notna(v) else 0.0)
            cell.border = BORDER
    ws.column_dimensions["A"].width = 14
    for j in range(2, len(pivot.columns) + 2):
        ws.column_dimensions[_col(j)].width = 14
    ws.freeze_panes = "B2"


def _write_value_pivot_sheet(wb: Workbook, res: ParseResult, value_col: str,
                             sheet_name: str, data_label: str):
    """通用"样本×物质"数值透视表（SN/峰高/半峰宽共用），结构与 RT表 一致。
    样本数据为空或该列全为 NaN 时跳过（不生成空表），兼容 AREA/RT 导出的实验。
    """
    sub = res.df[res.df["category"] == "sample"]
    if len(sub) == 0 or value_col not in sub.columns or sub[value_col].isna().all():
        return
    ws = wb.create_sheet(sheet_name)
    pivot = sub.pivot_table(index="sample", columns="analyte", values=value_col, aggfunc="first")

    _fmt_sk = getattr(res, "sample_format", "") or ""
    def _sort_key(s):
        g, ph, rp = parse_sample_name(s, pattern=_fmt_sk or None) or (99, "Z", 99)
        return (g, ph, rp)
    pivot = pivot.reindex(sorted(pivot.index, key=_sort_key))
    ws.cell(row=1, column=1, value="样本")
    for j, a in enumerate(pivot.columns, start=2):
        ws.cell(row=1, column=j, value=a)
    _style_header(ws, 1, len(pivot.columns) + 1)
    for i, (sample, row) in enumerate(pivot.iterrows(), start=2):
        ws.cell(row=i, column=1, value=sample).border = BORDER
        for j, a in enumerate(pivot.columns, start=2):
            v = row[a]
            cell = ws.cell(row=i, column=j, value=float(v) if pd.notna(v) else 0.0)
            cell.border = BORDER
    ws.column_dimensions["A"].width = 14
    for j in range(2, len(pivot.columns) + 2):
        ws.column_dimensions[_col(j)].width = 14
    ws.freeze_panes = "B2"
    ws.cell(row=len(pivot) + 3, column=1,
            value=f"注：{data_label}取自 ALLDATA 导出；空值记 0。").font = Font(italic=True, color="999999")


# ---------- 电压表 ----------

def write_voltage_sheet(wb: Workbook, voltage_df: Optional[pd.DataFrame] = None):
    """电压表：时间/电流(/来源)列。若提供 voltage_df 则写入，否则空表头等用户粘贴。
    voltage_df 列：时间(min) / 电流值 / 来源sheet（可选）
    """
    ws = wb.create_sheet("电压表")
    has_src = voltage_df is not None and len(voltage_df) and "来源sheet" in voltage_df.columns
    ws.cell(row=1, column=1, value="时间(min)")
    ws.cell(row=1, column=2, value="电流值")
    if has_src:
        ws.cell(row=1, column=3, value="来源sheet")
    _style_header(ws, 1, 3 if has_src else 2)
    if voltage_df is None or not len(voltage_df):
        ws.cell(row=2, column=1, value="（从此处粘贴时间-电流数据）").font = Font(italic=True, color="999999")
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 30
    if voltage_df is not None and len(voltage_df):
        for i, row in enumerate(voltage_df.itertuples(index=False), start=2):
            ws.cell(row=i, column=1, value=float(row[0]))
            ws.cell(row=i, column=2, value=float(row[1]))
            if has_src:
                ws.cell(row=i, column=3, value=row[2])
        # 冻结表头 + 数字格式
        ws.freeze_panes = "A2"
        for r in range(2, len(voltage_df) + 2):
            ws.cell(row=r, column=1).number_format = "0.0000"
            ws.cell(row=r, column=2).number_format = "0.0000"


# ---------- 主入口 ----------

def build_excel(
    out_path: str | Path,
    res: ParseResult,
    meta: dict,
    Va: float,
    Vd: float,
    use_sample: bool = True,
    voltage_df: Optional[pd.DataFrame] = None,
    excluded_samples: Optional[set] = None,
    gcond: Optional[pd.DataFrame] = None,
    qc_by_group: Optional[dict] = None,
    default_qc: Optional[dict] = None,
) -> Path:
    """
    生成单次实验的 Excel 留痕文件。

    参数
    ----
    out_path : 输出路径（如 .../20251112-DATA.xlsx）
    res      : parser 的 ParseResult
    meta     : 元数据字典
    Va, Vd   : 受体/供体相体积
    use_sample : 是否启用供体相
    voltage_df : 电压数据（时间/电流/来源sheet）DataFrame，可选
    excluded_samples : 被剔除的样本名集合（如 {'1-A-1'}）；M分表会标灰原始值并加修正列
    gcond    : 组级条件 DataFrame（含孔位/异常标签/备注等），写入总表组条件表 + M分表右侧
    qc_by_group : {组号: {role: 标准品名}}——各组使用的标准品组合（多标准品模式）
    default_qc  : {role: 标准品名}——未指定组合的组使用的默认组合
    """
    out_path = Path(out_path)
    excluded_samples = excluded_samples or set()
    wb = Workbook()
    default = wb.active
    wb.remove(default)

    # 计算指标（修正版用于全部数据表；不再另算一遍未用的原始版——白跑 O(物质×组)）
    metrics_fixed = compute_metrics(res.df, res.analytes, res.groups, Va, Vd,
                                    use_sample=use_sample, excluded_samples=excluded_samples,
                                    group_qc=qc_by_group, default_qc=default_qc)

    # 写各表
    write_summary_sheet(wb, res, metrics_fixed, meta, Va, Vd, use_sample, gcond=gcond,
                        default_qc=default_qc)
    write_data_summary_sheet(wb, metrics_fixed, gcond, res.analytes, res.groups,
                             sample_format=getattr(res, "sample_format", "") or "",
                             n_replicates=res.n_replicates, excluded_samples=excluded_samples)
    write_rt_sheet(wb, res)
    # SN / 峰高 / 半峰宽 原始数据表（仅当解析数据含这些列时生成，与网页端表格同步）
    _write_value_pivot_sheet(wb, res, "sn", "SN表", "信噪比")
    _write_value_pivot_sheet(wb, res, "height", "峰高表", "峰高")
    _write_value_pivot_sheet(wb, res, "width50", "半峰宽表", "半峰宽")
    write_voltage_sheet(wb, voltage_df)
    # 「组条件」sheet 已删除（与总表内组条件表完全重复）
    # 散点轴范围按整次实验（全部组）计算一次 → 所有 M 分表共用同一范围（各组可比）
    def _axis_base_lo(vals, base=0.0):
        _vs = [v for v in vals if v is not None]
        if not _vs or min(_vs) >= base:
            return base
        import math as _math
        return float(_math.floor(min(_vs) / 10.0) * 10.0)

    def _axis_base_hi(vals, base=100.0):
        _vs = [v for v in vals if v is not None]
        if not _vs or max(_vs) <= base:
            return base
        import math as _math
        return float(_math.ceil(max(_vs) / 10.0) * 10.0)

    _rec_all = [m.recovery for m in metrics_fixed if m.recovery is not None]
    _rsd_all = [m.A_rsd for m in metrics_fixed if m.A_rsd is not None]
    _shared_bounds = {"rec_lo": _axis_base_lo(_rec_all), "rec_hi": _axis_base_hi(_rec_all),
                      "rsd_lo": _axis_base_lo(_rsd_all), "rsd_hi": _axis_base_hi(_rsd_all)}
    for g in res.groups:
        write_group_sheet(wb, res, g, Va, Vd, use_sample, excluded_samples, gcond=gcond,
                          qc_cfg=(qc_by_group or {}).get(int(g)) or default_qc or {})
        try:
            _add_group_charts(wb[f"M{g}"], res, g, Va, Vd, use_sample, excluded_samples,
                              [m for m in metrics_fixed if m.group == g],
                              axis_bounds=_shared_bounds)
        except Exception:
            pass  # 图表生成失败不阻断数据文件（数据+公式始终完整）

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path


def _add_group_charts(ws, res: ParseResult, group: int, Va: float, Vd: float,
                      use_sample: bool, excluded_samples: Optional[set] = None,
                      metrics_rows: Optional[list] = None, axis_bounds: Optional[dict] = None):
    """M分表数据区下方（33行起）并排三张原生 Excel 图表：
    ① 三相分布（堆叠柱，仅启用供体相时）② 回收率-受体峰面积RSD 散点 ③ LogP-受体峰面积RSD 散点。
    图表引用的数据块写在图表下方（标注"图表数据"，可核对），数值为该组修正后结果。
    """
    from openpyxl.chart import BarChart, ScatterChart, Reference, Series
    from openpyxl.chart.marker import Marker

    n_a = len(res.analytes)
    if n_a == 0:
        return
    last_row = 1 + n_a
    chart_row = max(last_row + 2, 33)      # 数据表下方，33行起（用户指定 32 行以后）
    data_row = chart_row + 21              # 图高约 9.5cm ≈ 19 行，留余量
    mrow = {m.analyte: m for m in (metrics_rows or [])}

    def _logp(a):
        try:
            from . import reference as _ref
        except Exception:
            try:
                import reference as _ref
            except Exception:
                return None
        try:
            info = _ref.get_analyte_info(a)
            if info is not None:
                for lp_col in ("C-LogP", "Logp", "LogP"):
                    if lp_col in info.index:
                        v = _ref._try_float(info[lp_col])
                        if v is not None:
                            return v
        except Exception:
            pass
        return None

    # —— 数据块 1：三相分布（A列起）——
    dist_ok = False
    if use_sample:
        try:
            dist = compute_phase_distribution(res.df, res.analytes, [group], Va, Vd,
                                              excluded_samples=excluded_samples)
        except Exception:
            dist = []
        if dist:
            dist_ok = True
            ws.cell(row=data_row, column=1, value="物质")
            ws.cell(row=data_row, column=2, value="受体相%")
            ws.cell(row=data_row, column=3, value="膜%")
            ws.cell(row=data_row, column=4, value="供体相%")
            dist_map = {d["analyte"]: d for d in dist}
            for i, a in enumerate(res.analytes, start=1):
                rr_ = data_row + i
                ws.cell(row=rr_, column=1, value=a)
                d = dist_map.get(a)
                if d:
                    ws.cell(row=rr_, column=2, value=round(d["acceptor_pct"], 2))
                    ws.cell(row=rr_, column=3, value=round(d["membrane_pct"], 2))
                    ws.cell(row=rr_, column=4, value=round(d["donor_pct"], 2))

    # —— 数据块 2：回收率 / RSD（F列起）——
    ws.cell(row=data_row, column=6, value="物质")
    ws.cell(row=data_row, column=7, value="回收率%")
    ws.cell(row=data_row, column=8, value="受体峰面积RSD%")
    have_r = False
    for i, a in enumerate(res.analytes, start=1):
        rr_ = data_row + i
        ws.cell(row=rr_, column=6, value=a)
        m = mrow.get(a)
        if m is not None and m.recovery is not None:
            ws.cell(row=rr_, column=7, value=round(m.recovery, 3))
            have_r = True
        if m is not None and m.A_rsd is not None:
            ws.cell(row=rr_, column=8, value=round(m.A_rsd, 3))

    # —— 数据块 3：LogP / RSD（J列起）——
    ws.cell(row=data_row, column=10, value="物质")
    ws.cell(row=data_row, column=11, value="LogP")
    ws.cell(row=data_row, column=12, value="受体峰面积RSD%")
    have_lp = False
    for i, a in enumerate(res.analytes, start=1):
        rr_ = data_row + i
        ws.cell(row=rr_, column=10, value=a)
        lp = _logp(a)
        if lp is not None:
            ws.cell(row=rr_, column=11, value=lp)
            have_lp = True
        m = mrow.get(a)
        if m is not None and m.A_rsd is not None:
            ws.cell(row=rr_, column=12, value=round(m.A_rsd, 3))

    ws.cell(row=data_row - 2, column=1,
            value="—— 下方为上方图表引用的数据（本组修正后结果），可在此核对 ——").font = Font(
        italic=True, color="999999")

    n = n_a
    # —— 图 1：三相分布（堆叠柱状）。锚在 B 列：A 列是冻结列（freeze_panes=B2），
    # 锚在 A 的图表横向滑动时会被冻结区遮住左缘（用户反馈"截断"）——
    if dist_ok:
        bar = BarChart()
        bar.type = "col"
        bar.grouping = "stacked"
        bar.overlap = 100
        _d = Reference(ws, min_col=2, max_col=4, min_row=data_row, max_row=data_row + n)
        _c = Reference(ws, min_col=1, min_row=data_row + 1, max_row=data_row + n)
        bar.add_data(_d, titles_from_data=True)
        bar.set_categories(_c)
        # 类别必须写 strRef 且【带 strCache 缓存】：openpyxl 默认写 numRef（数字引用，
        # 物质名是文本 → 横坐标空白）；部分 Excel/WPS 对无缓存的文本引用也不显示
        from openpyxl.chart.data_source import AxDataSource, StrRef, StrData, StrVal
        _cat_cache = StrData(ptCount=n, pt=[StrVal(idx=i, v=str(a))
                                            for i, a in enumerate(res.analytes)])
        for _s_ in bar.ser:
            _s_.cat = AxDataSource(strRef=StrRef(f=str(_c), strCache=_cat_cache))
        bar.x_axis.delete = False
        bar.y_axis.delete = False
        bar.x_axis.tickLblPos = "nextTo"   # 缺省时部分查看器隐藏刻度标签
        bar.y_axis.tickLblPos = "nextTo"
        bar.title = f"M{group} 三相分布（受体相/膜/供体相，%）"
        bar.y_axis.title = "%"
        bar.x_axis.title = "物质"
        bar.legend.position = "b"
        bar.legend.overlay = False  # overlay=True 时图例会叠在绘图区上
        bar.width = 15
        bar.height = 9.5
        ws.add_chart(bar, f"B{chart_row}")

    _PALETTE = ["4E79A7", "F28E2B", "E15759", "76B7B2", "59A14F", "EDC948", "B07AA1",
                "FF9DA7", "9C755F", "BAB0AC", "1F77B4", "FF7F0E", "2CA02C", "D62728",
                "9467BD", "8C564B", "E377C2", "7F7F7F", "BCBD22", "17BECF"]

    def _scatter(x_col, y_col, title, x_title, y_title, symbol,
                 y_min=None, y_max=None, x_min=None, x_max=None):
        """每个药物一条序列（序列名=药物名 → 图例列出药物），单点散点+区分色。
        x_col/y_col 为数据块中的列号；任一缺失的药物不建序列。
        y_min/y_max、x_min/x_max：固定轴范围（各组之间可直接对比）。""" 
        from openpyxl.chart.shapes import GraphicalProperties
        sc = ScatterChart()
        sc.style = 13
        sc.title = title
        sc.x_axis.title = x_title
        sc.y_axis.title = y_title
        sc.x_axis.delete = False
        sc.y_axis.delete = False
        sc.x_axis.tickLblPos = "nextTo"
        sc.y_axis.tickLblPos = "nextTo"
        if y_min is not None and y_max is not None:
            sc.y_axis.scaling.min = y_min
            sc.y_axis.scaling.max = y_max
        if x_min is not None and x_max is not None:
            sc.x_axis.scaling.min = x_min
            sc.x_axis.scaling.max = x_max
        n_series = 0
        for i, a in enumerate(res.analytes):
            rr_ = data_row + 1 + i
            if (ws.cell(row=rr_, column=x_col).value is None
                    or ws.cell(row=rr_, column=y_col).value is None):
                continue  # X 或 Y 缺失（如 LogP 查不到）→ 不建序列
            _x = Reference(ws, min_col=x_col, min_row=rr_, max_row=rr_)
            _y = Reference(ws, min_col=y_col, min_row=rr_, max_row=rr_)
            s = Series(_y, _x, title=str(a))  # 显式药物名（title_from_data 在散点上会显示成数值）
            color = _PALETTE[n_series % len(_PALETTE)]
            s.marker = Marker(symbol=symbol, size=6,
                              spPr=GraphicalProperties(solidFill=color))
            s.graphicalProperties.line.noFill = True
            sc.series.append(s)
            n_series += 1
        sc.legend.position = "r"
        sc.legend.overlay = False  # overlay=True 时图例会叠在绘图区上
        sc.width = 15
        sc.height = 9.5
        return sc

    # —— 图 2 / 图 3：散点。图2 横=受体峰面积RSD% 纵=回收率%；图3 横=LogP 纵=回收率%。
    #    轴范围 = 0~100 为主、自动扩展容纳超出点（扩展到 10 的整倍数，如 0~120），
    #    各组之间仍可直接对比。与图1 水平等距（B / J / R，避开冻结列）——
    _ab = axis_bounds or {}
    _rec_lo, _rec_hi = _ab.get("rec_lo", 0.0), _ab.get("rec_hi", 100.0)
    _rsd_lo, _rsd_hi = _ab.get("rsd_lo", 0.0), _ab.get("rsd_hi", 100.0)
    if have_r:
        ws.add_chart(_scatter(8, 7, f"M{group} 受体峰面积RSD-回收率 散点",
                              "受体峰面积RSD%", "回收率%", "circle",
                              y_min=_rec_lo, y_max=_rec_hi,
                              x_min=_rsd_lo, x_max=_rsd_hi),
                     f"J{chart_row}")
    if have_lp and have_r:
        ws.add_chart(_scatter(11, 7, f"M{group} LogP-回收率 散点",
                              "LogP", "回收率%", "diamond",
                              y_min=_rec_lo, y_max=_rec_hi),
                     f"R{chart_row}")


def write_data_summary_sheet(wb: Workbook, metrics: list, gcond: Optional[pd.DataFrame],
                             analytes: list, groups: list,
                             n_replicates: int = 3, excluded_samples: Optional[set] = None,
                             sample_format: str = ""):
    """「数据汇总」sheet：物质×实验组矩阵（类似手工作表样式）。
    依次排列：物质性质列（LogP/电荷@供体pH）+ 每组一列；五个矩阵块：
    回收率 / 供体相至SLM转移率 / SLM至受体相转移率 / 受体峰面积RSD / 供体峰面积RSD。
    每块带色阶条件格式 + 底部统计行（有效数 / >=40 / 20~40 / <20）。
    """
    import json as _json
    _fmt_eff = sample_format  # 实验实际生效的样品名格式（剔除计数解析用）
    ws = wb.create_sheet("数据汇总")
    # 组信息：{grp: (组别名, pH_donor, slm/条件简述)}
    grp_info = {}
    if gcond is not None and len(gcond):
        for _, r in gcond.iterrows():
            if r["grp"] == 0:
                continue
            grp_info[int(r["grp"])] = (r.get("group_label") or f"M{int(r['grp'])}",
                                       r.get("pH_donor"), r.get("slm_main") or "")
    # 物质性质（LogP / 电荷@供体相pH）
    prop_cache = {}

    def _props(analyte):
        if analyte in prop_cache:
            return prop_cache[analyte]
        logp = charge = None
        try:
            from . import reference as _ref
        except Exception:
            try:
                import reference as _ref
            except Exception:
                _ref = None
        if _ref is not None:
            try:
                info = _ref.get_analyte_info(analyte)
                if info is not None:
                    for lp_col in ("C-LogP", "Logp", "LogP"):  # 新旧表列名兼容
                        if lp_col in info.index:
                            logp = _ref._try_float(info[lp_col])
                            if logp is not None:
                                break
                ph = None
                for gi in groups:
                    if gi in grp_info and grp_info[gi][1] not in (None, ""):
                        ph = float(grp_info[gi][1]); break
                if ph is not None:
                    charge = _ref.get_property_by_ph(analyte, "charge", ph)
            except Exception:
                pass
        prop_cache[analyte] = (logp, charge)
        return logp, charge

    # 物质排序：按 LogP 升序（与手工表一致），无 LogP 的排最后
    analytes_sorted = sorted(analytes, key=lambda a: (_props(a)[0] is None, _props(a)[0] if _props(a)[0] is not None else 0))

    blocks = [
        ("recovery", "回收率（物质 × 实验组，%）", (0, 100), "rate"),
        ("dm", "供体相至SLM转移率（物质 × 实验组，%）", (0, 100), "rate"),
        ("ma", "SLM至受体相转移率（物质 × 实验组，%）", (0, 100), "rate"),
        ("A_rsd", "受体峰面积RSD（物质 × 实验组，%）", (0, 50), "rsd"),
        ("S_rsd", "供体峰面积RSD（物质 × 实验组，%）", (0, 50), "rsd"),
    ]
    n_prop_cols = 2
    row = 1
    # 每块宽度 = 2性质列 + N组列
    for col_idx, (field, title, rng, kind) in enumerate(blocks):
        # 标题行
        ws.cell(row=row, column=1, value=title).font = Font(bold=True, size=12)
        row += 1
        if not groups:
            # 无样品组（如纯质控数据）：跳过矩阵块（空组的反向区域会让条件格式崩溃）
            continue
        # 组头两行：M{n} / 组别名
        ws.cell(row=row, column=1, value="物质").font = Font(bold=True)
        ws.cell(row=row, column=2, value="LogP").font = Font(bold=True)
        ws.cell(row=row, column=3, value="charge@供体pH").font = Font(bold=True)
        for gi, g in enumerate(groups):
            lbl, ph, slm = grp_info.get(g, (f"M{g}", None, ""))
            ws.cell(row=row, column=4 + gi, value=f"M{g}").font = Font(bold=True)
            _lbl_cell = ws.cell(row=row + 1, column=4 + gi, value=lbl)  # 全名写入，不截断
            _lbl_cell.alignment = CENTER  # 长名自动换行显示
        ws.cell(row=row + 1, column=1, value="（按LogP升序）").font = Font(italic=True, size=9)
        header_row = row + 1
        # 数据行
        data_start = row + 2
        from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
        for ai, analyte in enumerate(analytes_sorted):
            rr = data_start + ai
            ws.cell(row=rr, column=1, value=analyte)
            logp, charge = _props(analyte)
            if logp is not None:
                ws.cell(row=rr, column=2, value=round(logp, 3))
            if charge is not None:
                ws.cell(row=rr, column=3, value=round(charge, 3))
        # 值填入 + 条件格式
        value_by = {}
        for m in metrics:
            value_by[(m.analyte, m.group)] = getattr(m, field, None)
        first_data_col = 4
        last_col = 3 + len(groups)
        for gi, g in enumerate(groups):
            for ai, analyte in enumerate(analytes_sorted):
                rr = data_start + ai
                v = value_by.get((analyte, g))
                if v is not None:
                    ws.cell(row=rr, column=first_data_col + gi, value=round(v, 3))
        col_l = _col(first_data_col)
        col_r = _col(last_col)
        rng_str = f"{col_l}{data_start}:{col_r}{data_start + len(analytes_sorted) - 1}"
        if kind == "rate":
            ws.conditional_formatting.add(rng_str, ColorScaleRule(
                start_type="num", start_value=rng[0], start_color="F8696B",
                mid_type="num", mid_value=(rng[0] + rng[1]) / 2, mid_color="FFEB84",
                end_type="num", end_value=rng[1], end_color="63BE7B"))
        else:  # rsd：>15 红
            ws.conditional_formatting.add(rng_str, ColorScaleRule(
                start_type="num", start_value=0, start_color="FFFFFF",
                end_type="num", end_value=30, end_color="F8696B"))
        # 统计行（全部用 Excel 公式，便于核对）
        stat_row = data_start + len(analytes_sorted)
        stat_defs = [
            ("有效数", None),  # 特殊：重复数-该组被剔除数（公式）
            ("R>=40", 'COUNTIF({rng},">=40")'),
            ("20~40", 'COUNTIFS({rng},">=20",{rng},"<40")'),
            ("<20", 'COUNTIF({rng},"<20")'),
        ]
        for si, (lbl, tpl) in enumerate(stat_defs):
            r2 = stat_row + si
            ws.cell(row=r2, column=1, value=lbl).font = Font(bold=si == 0)
            for gi, g in enumerate(groups):
                cl = _col(first_data_col + gi)
                cell_rng = f"{cl}{data_start}:{cl}{data_start + len(analytes_sorted) - 1}"
                if lbl == "有效数":
                    # 该组被剔除的样本数（受体+供体合计）——按解析取组号，
                    # 不得按 "组号-" 前缀猜测（A1-1 等自定义格式下恒为 0）
                    n_excl = sum(1 for s in (excluded_samples or set())
                                 if (parse_sample_name(str(s), pattern=(_fmt_eff or None))
                                       or (None,))[0] == g)
                    ws.cell(row=r2, column=first_data_col + gi,
                            value=f"={n_replicates}-{n_excl}")
                else:
                    ws.cell(row=r2, column=first_data_col + gi,
                            value="=" + tpl.format(rng=cell_rng))
        row = stat_row + len(stat_defs) + 2  # 空两行进入下一块

    # 列宽
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 14
    for gi in range(len(groups)):
        ws.column_dimensions[_col(4 + gi)].width = 16
    ws.freeze_panes = "D1"


# 「组条件」sheet 已移除：与总表内的组条件表完全重复（write_summary_sheet 负责写入总表）
