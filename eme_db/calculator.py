# -*- coding: utf-8 -*-
"""
EME 实验数据管理系统 - 计算引擎

职责：
1. STD/BLK 聚合：按物质取峰面积平均值 + RSD（多次进样/多个同名样）
2. 样品组聚合：每组的 3 次重复（A 相/S 相各一组）取均值 + RSD
3. 回收率（空白双轨 + 校正系数）
4. 转移率 D→M / M→A（超 0~100 标红）
5. 数据剔除：剔除指定样本/STD 行后重算（留痕）
6. 表达式引擎：simpleeval 解析用户自定义公式

关键设计：
- 所有"均值"都是简单算术平均；RSD = std/mean×100%（百分比）
- 空白用对应相的空白（受体空白对应受体相，供体空白对应供体相）
- 校正系数 corr = 供体标准品均值 / 受体标准品均值；两相溶液相同时，用户填同名 → corr=1
"""
from __future__ import annotations
import math
from dataclasses import dataclass, field
from typing import Optional

import numpy as np
import pandas as pd

try:
    from . import config
except ImportError:
    import config


# ---------- 基础统计 ----------

def mean(values) -> float:
    """算术平均，空列表/全0返回0"""
    v = [x for x in values if x is not None and not (isinstance(x, float) and math.isnan(x))]
    if not v:
        return 0.0
    return float(np.mean(v))


def rsd_percent(values) -> float:
    """RSD% = std/mean×100；mean=0 或样本数<2 返回 0"""
    v = [x for x in values if x is not None and not (isinstance(x, float) and math.isnan(x))]
    if len(v) < 2:
        return 0.0
    m = float(np.mean(v))
    if m == 0:
        return 0.0
    return float(np.std(v, ddof=1) / m * 100.0)


# ---------- STD/BLK 聚合 ----------

def aggregate_qc(df: pd.DataFrame, category: str, analyte: str,
                 value_col: str = "area", sample_name: str = None) -> tuple[float, float, list[float]]:
    """聚合某类质控/空白（如 std_acceptor）某物质的指定值列（area/height/width50/rt...）。
    sample_name: 多标准品组合时按标准品名过滤（None=该类别全部行，兼容单标准品）。
    返回 (均值, RSD%, [原始值列表])
    """
    sub = df[(df["category"] == category) & (df["analyte"] == analyte)]
    if sample_name:
        sub = sub[sub["sample"] == sample_name]
    vals = sub[value_col].tolist()
    return mean(vals), rsd_percent(vals), vals


# ---------- 样品组聚合 ----------

def aggregate_sample_phase(df: pd.DataFrame, group: int, phase: str, analyte: str,
                           excluded_samples: Optional[set] = None) -> tuple[float, float, list[float]]:
    """聚合某组某相某物质的 3 次重复。
    excluded_samples: 要剔除的样本名集合（如 {'1-A-1'}），这些样本不计入
    返回 (均值, RSD%, [原始值列表])
    """
    excluded_samples = excluded_samples or set()
    mask = (
        (df["category"] == "sample") &
        (df["group"] == group) &
        (df["phase"] == phase) &
        (df["analyte"] == analyte)
    )
    sub = df[mask & ~df["sample"].isin(excluded_samples)]
    vals = sub["area"].tolist()
    return mean(vals), rsd_percent(vals), vals


# ---------- 核心计算：回收率 + 转移率 ----------

@dataclass
class MetricRow:
    """单个物质 × 单个样品组的计算结果行"""
    group: int
    analyte: str
    # 原始聚合
    A_mean: float                  # 受体相峰面积均值
    A_rsd: float                   # 受体相 RSD%
    S_mean: float                  # 供体相峰面积均值
    S_rsd: float                   # 供体相 RSD%
    # 质控参考
    STD_acceptor_mean: float
    STD_donor_mean: float
    blk_acceptor_mean: float
    blk_donor_mean: float
    corr: float                    # 校正系数
    # 计算结果
    recovery: Optional[float] = None          # 回收率%
    dm: Optional[float] = None                # D→M 转移率%
    ma: Optional[float] = None                # M→A 转移率%
    recovery_flag: str = ""                   # "" / "high_rsd"
    dm_flag: str = ""                         # "" / "out_of_range"
    ma_flag: str = ""                         # "" / "out_of_range" / "no_donor_data"


def compute_metrics(
    df: pd.DataFrame,
    analytes: list[str],
    groups: list[int],
    Va: float,
    Vd: float,
    use_sample: bool = True,
    excluded_samples: Optional[set] = None,
    rsd_threshold: float = None,
    group_qc: Optional[dict] = None,
    default_qc: Optional[dict] = None,
) -> list[MetricRow]:
    """
    计算每个物质 × 每个样品组的回收率、转移率、RSD。

    参数
    ----
    df : parser 输出的干净长表
    analytes, groups : 物质列表、样品组列表
    Va, Vd : 受体/供体相体积 μL
    use_sample : 是否启用供体相（转移率计算）
    excluded_samples : 剔除的样本名集合
    rsd_threshold : RSD 阈值%，超过标记 high_rsd
    group_qc : {组号: {role: 标准品名}}——各组使用不同的供受体相标准品组合
    default_qc : {role: 标准品名}——未指定组合的组使用的默认组合
    """
    excluded_samples = excluded_samples or set()
    rsd_threshold = rsd_threshold if rsd_threshold is not None else config.RSD_THRESHOLD
    group_qc = group_qc or {}
    default_qc = default_qc or {}
    results: list[MetricRow] = []

    for analyte in analytes:
        for g in groups:
            # 各组标准品组合：该组指定 > 默认组合；均无时该类别全部行（兼容单标准品）
            _qc_cfg = group_qc.get(int(g)) or default_qc
            std_a_mean, std_a_rsd, _ = aggregate_qc(df, "std_acceptor", analyte, sample_name=_qc_cfg.get("std_acceptor"))
            std_d_mean, std_d_rsd, _ = aggregate_qc(df, "std_donor", analyte, sample_name=_qc_cfg.get("std_donor"))
            blk_a_mean, _, _ = aggregate_qc(df, "blk_acceptor", analyte, sample_name=_qc_cfg.get("blk_acceptor"))
            blk_d_mean, _, _ = aggregate_qc(df, "blk_donor", analyte, sample_name=_qc_cfg.get("blk_donor"))

            # 校正系数：供体标准品/受体标准品；分母为0时取1
            corr = (std_d_mean / std_a_mean) if std_a_mean != 0 else 1.0
            A_mean, A_rsd, _ = aggregate_sample_phase(df, g, "A", analyte, excluded_samples)
            if use_sample:
                S_mean, S_rsd, _ = aggregate_sample_phase(df, g, "S", analyte, excluded_samples)
            else:
                # 未启用供体相：不计算供体相RSD（即使数据中有S样本）
                S_mean, S_rsd = 0.0, None

            row = MetricRow(
                group=g, analyte=analyte,
                A_mean=A_mean, A_rsd=A_rsd,
                S_mean=S_mean, S_rsd=S_rsd,
                STD_acceptor_mean=std_a_mean, STD_donor_mean=std_d_mean,
                blk_acceptor_mean=blk_a_mean, blk_donor_mean=blk_d_mean,
                corr=corr,
            )

            # —— 回收率 ——
            # ((受体峰面积 − 受体空白均值) × Va) / ((供体标准品均值 − 供体空白均值) × Vd) × corr
            denom = (std_d_mean - blk_d_mean) * Vd
            numer = (A_mean - blk_a_mean) * Va
            if denom != 0:
                row.recovery = numer / denom * corr * 100.0
            else:
                row.recovery = None

            # RSD 标记（受体或供体 RSD 超阈；供体相未启用时 S_rsd 为 None 不参与）
            if A_rsd > rsd_threshold or (S_rsd is not None and S_rsd > rsd_threshold):
                row.recovery_flag = "high_rsd"

            # —— 转移率（仅 use_sample 且有供体相数据时）——
            if use_sample and S_mean > 0:
                # D→M = 100% − (供体均值−供体空白)/(供体标准品均值−供体空白)
                denom_dm = std_d_mean - blk_d_mean
                if denom_dm != 0:
                    dm = 100.0 - (S_mean - blk_d_mean) / denom_dm * 100.0
                    row.dm = dm
                    if dm < config.TRANSFER_RATE_MIN or dm > config.TRANSFER_RATE_MAX:
                        row.dm_flag = "out_of_range"
                else:
                    row.dm = None

                # M→A = 100% × 回收率 / D→M
                if row.dm is not None and row.dm != 0 and row.recovery is not None:
                    ma = 100.0 * row.recovery / row.dm
                    row.ma = ma
                    if ma < config.TRANSFER_RATE_MIN or ma > config.TRANSFER_RATE_MAX:
                        row.ma_flag = "out_of_range"
                else:
                    row.ma = None
            else:
                row.ma_flag = "no_donor_data"

            results.append(row)

    return results


def metrics_to_dataframe(metrics: list[MetricRow]) -> pd.DataFrame:
    """把计算结果转成 DataFrame（全称命名，方便展示/导出）"""
    rows = []
    for m in metrics:
        rows.append({
            "组号": m.group,
            "物质": m.analyte,
            "受体峰面积平均值": m.A_mean,
            "受体峰面积RSD%": round(m.A_rsd, 2) if m.A_rsd is not None else None,
            "供体峰面积平均值": m.S_mean,
            "供体峰面积RSD%": round(m.S_rsd, 2) if m.S_rsd is not None else None,
            "回收率%": round(m.recovery, 2) if m.recovery is not None else None,
            "供体相至SLM转移率%": round(m.dm, 2) if m.dm is not None else None,
            "SLM至受体相转移率%": round(m.ma, 2) if m.ma is not None else None,
            "校正系数": round(m.corr, 4),
        })
    return pd.DataFrame(rows)


# ---------- RT 统计（每组样品的保留时间，不计算RSD，仅用于出图）----------

def collect_rt(df: pd.DataFrame, analytes: list[str], groups: list[int]) -> pd.DataFrame:
    """收集各样本的保留时间（长表），用于画散点图（X=样本数，Y=RT）"""
    sub = df[df["category"] == "sample"][["sample", "group", "phase", "replicate", "analyte", "rt"]].copy()
    return sub.reset_index(drop=True)


def get_std_injection_matrix(df: pd.DataFrame, std_category: str = "std_acceptor",
                             analytes: list[str] = None) -> pd.DataFrame:
    """
    提取标准品"每次进样 × 物质"的峰面积矩阵。
    同名标准品（如 Accept-STD-100ppb-1）多次进样，按出现顺序编为 #1,#2,...
    返回 DataFrame：行=进样针(#1..#N)，列=各物质峰面积，index 名为"进样针"。
    """
    sub = df[df["category"] == std_category].copy()
    if len(sub) == 0:
        return pd.DataFrame()
    if analytes is None:
        analytes = sorted(sub["analyte"].unique())
    # 按 sample+analyte 的出现顺序给每次进样编号（injection_idx 已在 parser 算好）
    # 但同名 sample 多次进样需要跨 analyte 统一编号。
    # 用 (sample, injection_idx) 标识一次"针"；injection_idx 在 parser 是按 (sample,component) cumcount
    # 为统一，这里按 sample 分组、取每组 injection_idx 的唯一值作为针号
    if "injection_idx" not in sub.columns:
        return pd.DataFrame()
    # 每个 (sample, injection_idx) = 一针
    sub["针号"] = sub.groupby("sample")["injection_idx"].rank(method="dense").astype(int)
    # pivot：行=针号，列=物质
    # 但多个 sample 会冲突；取第一个 sample（通常一次实验一个标准品名）
    samples = sorted(sub["sample"].unique())
    frames = []
    for s in samples:
        ss = sub[sub["sample"] == s]
        pivot = ss.pivot_table(index="针号", columns="analyte", values="area", aggfunc="first")
        pivot.index = [f"{s} #{i}" for i in pivot.index]
        frames.append(pivot)
    if not frames:
        return pd.DataFrame()
    out = pd.concat(frames)
    # 只保留指定物质列
    out = out.reindex(columns=[a for a in analytes if a in out.columns])
    return out


def get_sample_phase_matrix(df: pd.DataFrame, group: int, phase: str,
                            analytes: list[str] = None) -> pd.DataFrame:
    """
    提取某组某相的"物质 × 重复"峰面积矩阵 + RSD。
    返回 DataFrame：行=物质，列=重复1/2/3峰面积 + 均值 + RSD%。
    """
    sub = df[(df["category"] == "sample") & (df["group"] == group) & (df["phase"] == phase)].copy()
    if len(sub) == 0:
        return pd.DataFrame()
    if analytes is None:
        analytes = sorted(sub["analyte"].unique())
    # pivot：行=物质，列=重复号
    pivot = sub.pivot_table(index="analyte", columns="replicate", values="area", aggfunc="first")
    pivot = pivot.reindex([a for a in analytes if a in pivot.index])
    # 重命名列为 1/2/3
    pivot.columns = [f"重复{int(c)}" for c in pivot.columns]
    # 均值/RSD
    num_cols = pivot.select_dtypes(include="number").columns.tolist()
    pivot["平均值"] = pivot[num_cols].mean(axis=1)
    pivot["RSD%"] = pivot[num_cols].std(axis=1, ddof=1) / pivot["平均值"] * 100
    pivot["RSD%"] = pivot["RSD%"].where(pivot["平均值"] != 0, 0)
    return pivot.reset_index()


# ---------- 三厢分布（供体相/膜/受体相，百分比堆积图数据）----------

def compute_phase_distribution(df: pd.DataFrame, analytes: list, groups: list,
                               Va: float, Vd: float,
                               excluded_samples: Optional[set] = None,
                               group_qc: Optional[dict] = None,
                               default_qc: Optional[dict] = None) -> list[dict]:
    """
    计算各实验组各物质在 供体相/膜/受体相 的量与百分比。
    - 供体相量 = (供体峰面积均值 − 供体空白) × Vd
    - 受体相量 = (受体峰面积均值 − 受体空白) × Va
    - 膜相量   = (供体标准品均值 − 供体空白) × Vd − 受体相量 × 校正系数；<0 记 0
    - 百分比 = 各相量 / 三相总量 × 100
    返回 [{analyte, group, donor_amt, membrane_amt, acceptor_amt,
           donor_pct, membrane_pct, acceptor_pct, donor_rsd, acceptor_rsd}]
    """
    excluded_samples = excluded_samples or set()
    group_qc = group_qc or {}
    default_qc = default_qc or {}
    out = []
    for analyte in analytes:
        for g in groups:
            _qc_cfg = group_qc.get(int(g)) or default_qc or {}
            std_a_mean, _, _ = aggregate_qc(df, "std_acceptor", analyte, sample_name=_qc_cfg.get("std_acceptor"))
            std_d_mean, _, _ = aggregate_qc(df, "std_donor", analyte, sample_name=_qc_cfg.get("std_donor"))
            blk_a_mean, _, _ = aggregate_qc(df, "blk_acceptor", analyte, sample_name=_qc_cfg.get("blk_acceptor"))
            blk_d_mean, _, _ = aggregate_qc(df, "blk_donor", analyte, sample_name=_qc_cfg.get("blk_donor"))
            corr = (std_d_mean / std_a_mean) if std_a_mean else 1.0
            total0 = (std_d_mean - blk_d_mean) * Vd  # 初始总量（供体相标准品折算）
            if total0 <= 0:
                continue
            A_mean, A_rsd, _ = aggregate_sample_phase(df, g, "A", analyte, excluded_samples)
            S_mean, S_rsd, _ = aggregate_sample_phase(df, g, "S", analyte, excluded_samples)
            acceptor_amt = max(0.0, (A_mean - blk_a_mean) * Va)
            donor_amt = max(0.0, (S_mean - blk_d_mean) * Vd)
            membrane_amt = max(0.0, total0 - acceptor_amt * corr)
            tot = donor_amt + membrane_amt + acceptor_amt
            if tot <= 0:
                continue
            out.append({
                "analyte": analyte, "group": g,
                "donor_amt": donor_amt, "membrane_amt": membrane_amt, "acceptor_amt": acceptor_amt,
                "donor_pct": donor_amt / tot * 100,
                "membrane_pct": membrane_amt / tot * 100,
                "acceptor_pct": acceptor_amt / tot * 100,
                "donor_rsd": S_rsd, "acceptor_rsd": A_rsd,
            })
    return out


# ---------- 表达式引擎（自定义公式）----------

def eval_custom_formula(
    expr: str,
    variables: dict,
) -> tuple[Optional[float], Optional[str]]:
    """
    安全解析用户自定义公式表达式。
    返回 (结果, 错误信息)；出错时结果为 None。
    """
    if not expr or not expr.strip():
        return None, "公式为空"
    try:
        from simpleeval import simple_eval, EvalWithCompoundTypes
        # 用 EvalWithCompoundTypes 允许 if-else 等
        evaluator = EvalWithCompoundTypes(names=variables)
        val = evaluator.eval(expr)
        if isinstance(val, bool):
            return float(val), None
        if isinstance(val, (int, float)):
            f = float(val)
            if f != f or f in (float("inf"), float("-inf")):  # NaN/inf 不入库不入图
                return None, "公式结果为 NaN/无穷大（检查除零）"
            return f, None
        return None, f"公式结果不是数值: {type(val)}"
    except Exception as e:
        return None, f"公式解析错误: {e}"


def build_formula_variables(metrics_row: MetricRow, Va: float, Vd: float,
                            pH_donor: Optional[float] = None,
                            pH_acceptor: Optional[float] = None) -> dict:
    """为某行计算结果构造公式变量命名空间"""
    m = metrics_row
    return {
        "A_mean": m.A_mean,
        "S_mean": m.S_mean,
        "STD_acceptor_mean": m.STD_acceptor_mean,
        "STD_donor_mean": m.STD_donor_mean,
        "blk_acceptor_mean": m.blk_acceptor_mean,
        "blk_donor_mean": m.blk_donor_mean,
        "Va": Va,
        "Vd": Vd,
        "pH_donor": pH_donor if pH_donor is not None else 0.0,
        "pH_acceptor": pH_acceptor if pH_acceptor is not None else 0.0,
        "corr": m.corr,
    }


# ---------- 单次实验汇总（给跨实验看板用）----------

@dataclass
class ExperimentSummary:
    """一次实验的汇总（各物质的平均回收率/转移率 + RSD）"""
    analyte: str
    recovery_mean: float        # 各样品组回收率的平均
    recovery_rsd: float         # 各样品组回收率的 RSD%
    dm_mean: float
    dm_rsd: float
    ma_mean: float
    ma_rsd: float
    std_rsd: float              # 标准品峰面积 RSD%（跨多次进样）


def summarize_experiment(metrics: list[MetricRow], rsd_threshold: float = None) -> list[ExperimentSummary]:
    """按物质汇总一次实验：把各样品组的结果再平均"""
    rsd_threshold = rsd_threshold if rsd_threshold is not None else config.RSD_THRESHOLD
    by_analyte: dict[str, list[MetricRow]] = {}
    for m in metrics:
        by_analyte.setdefault(m.analyte, []).append(m)

    out = []
    for analyte, rows in by_analyte.items():
        recs = [r.recovery for r in rows if r.recovery is not None]
        dms = [r.dm for r in rows if r.dm is not None]
        mas = [r.ma for r in rows if r.ma is not None]
        stds = [r.STD_donor_mean for r in rows]  # 标准品均值在同物质各组相同，取来算RSD不合适
        # 标准品 RSD 应来自原始多次进样；这里用各组 STD_donor_mean 的离散度近似（通常相同）
        # 真实的标准品RSD在 parser→calculator 的 aggregate_qc 中算出，单独传入更准
        out.append(ExperimentSummary(
            analyte=analyte,
            recovery_mean=mean(recs),
            recovery_rsd=rsd_percent(recs),
            dm_mean=mean(dms),
            dm_rsd=rsd_percent(dms),
            ma_mean=mean(mas),
            ma_rsd=rsd_percent(mas),
            std_rsd=0.0,  # 由调用方单独填入（见 get_std_rsd_by_analyte）
        ))
    return out


def get_std_rsd_by_analyte(df: pd.DataFrame, analytes: list[str]) -> dict[str, float]:
    """每个物质的标准品（受体相标准品）峰面积跨多次进样的 RSD%"""
    out = {}
    for a in analytes:
        _, rsd, _ = aggregate_qc(df, "std_acceptor", a)
        out[a] = rsd
    return out


# ---------- 命令行自检 ----------
if __name__ == "__main__":
    import sys
    sys.path.insert(0, ".")
    import parser as P

    fld = sys.argv[1] if len(sys.argv) > 1 else "../data/experiments/20251112-8M"
    res = P.parse_experiment(fld)

    print("=" * 70)
    print("计算引擎自检")
    print(f"物质数: {len(res.analytes)}, 样品组: {res.groups}")
    print()

    metrics = compute_metrics(
        res.df, res.analytes, res.groups,
        Va=config.DEFAULT_VA, Vd=config.DEFAULT_VD,
        use_sample=True,
    )

    # 取前 3 个物质 × 组1 展示
    show = [m for m in metrics if m.group == 1][:3]
    for m in show:
        print(f"[{m.analyte} 组{m.group}]")
        print(f"  A_mean={m.A_mean:.2e} (RSD={m.A_rsd:.1f}%)  S_mean={m.S_mean:.2e} (RSD={m.S_rsd:.1f}%)")
        print(f"  STD_donor={m.STD_donor_mean:.2e} blk_donor={m.blk_donor_mean:.2e}")
        print(f"  corr={m.corr:.4f}")
        print(f"  回收率={m.recovery:.2f}%  D→M={m.dm:.2f}%{'⚠' if m.dm_flag else ''}  M→A={m.ma:.2f}%{'⚠' if m.ma_flag else ''}")
        print()

    # 测试自定义公式
    print("===== 自定义公式测试 =====")
    m0 = metrics[0]
    vs = build_formula_variables(m0, config.DEFAULT_VA, config.DEFAULT_VD, pH_donor=7.0, pH_acceptor=3.0)
    print("变量:", {k: (f'{v:.3e}' if isinstance(v, float) and abs(v)>1e4 else v) for k, v in vs.items()})
    expr = "(A_mean - blk_acceptor_mean) * Va / ((STD_donor_mean - blk_donor_mean) * Vd) * corr * 100"
    val, err = eval_custom_formula(expr, vs)
    print(f"公式: {expr}")
    print(f"结果: {val:.4f}%  (应与回收率 {m0.recovery:.4f}% 一致)" if val else f"错误: {err}")
