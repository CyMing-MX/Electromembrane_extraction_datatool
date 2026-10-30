# -*- coding: utf-8 -*-
"""
EME 实验数据管理系统 - 跨实验看板（Plotly 交互图）

提供：
1. 质控双轨图：标准品峰面积 vs 日期（固定±15%质控线 + 实际跨实验RSD值）；RT同理
2. 单次实验 RSD 表：数字表格，>15%标红
3. 自定义看板：X/Y/颜色/筛选自由组合
4. 物质下钻：点击物质看历次趋势
"""
from __future__ import annotations
from typing import Optional

import pandas as pd
import plotly.graph_objects as go
import plotly.express as px

try:
    from . import config
except ImportError:
    import config


# ---------- 质控双轨图（标准品峰面积 vs 日期）----------

# ---------- 质控双轨图（标准品峰面积/RT/半峰宽 vs 日期，支持多物质叠加）----------

_QC_COLORS = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
              "#8c564b", "#e377c2", "#17becf", "#bcbd22", "#7f7f7f"]


def _qc_overlay_chart(qc_df: pd.DataFrame, analytes: list, value_col: str,
                      value_name: str, fmt: str) -> go.Figure:
    """多物质叠加质控图：每物质一条曲线（专属颜色）+ 该物质的均值/±15% 参考线（同色虚线）。"""
    fig = go.Figure()
    plotted = False
    rtxt = []
    for i, analyte in enumerate(analytes):
        d = qc_df[qc_df["analyte"] == analyte].dropna(subset=[value_col])
        d = d[d[value_col] > 0].sort_values("exp_date")
        if len(d) == 0:
            continue
        plotted = True
        color = _QC_COLORS[i % len(_QC_COLORS)]
        mean_val = d[value_col].mean()
        upper, lower = mean_val * 1.15, mean_val * 0.85
        actual_rsd = d[value_col].std(ddof=1) / mean_val * 100 if len(d) > 1 and mean_val else 0
        rtxt.append(f"{analyte}: RSD={actual_rsd:.1f}%")
        x_all = d["exp_date"].tolist()
        fig.add_trace(go.Scatter(x=x_all, y=d[value_col], mode="lines+markers",
                                 name=analyte, line=dict(color=color, width=2),
                                 marker=dict(size=8),
                                 hovertemplate=f"{analyte}<br>日期=%{{x}}<br>{value_name}=%{{y:{fmt}}}<extra></extra>"))
        fig.add_trace(go.Scatter(x=x_all, y=d[value_col] * 1.15, mode="lines",
                                 line=dict(color=color, width=1, dash="dot"), opacity=0.5,
                                 showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=x_all, y=d[value_col] * 0.85, mode="lines",
                                 line=dict(color=color, width=1, dash="dot"), opacity=0.5,
                                 showlegend=False, hoverinfo="skip"))
        if len(analytes) == 1:
            fig.add_hline(y=upper, line_dash="dash", line_color="red", opacity=0.6)
            fig.add_hline(y=lower, line_dash="dash", line_color="red", opacity=0.6)
            fig.add_hline(y=mean_val, line_dash="dot", line_color="green", opacity=0.8)
    if not plotted:
        fig.update_layout(title="无质控数据", template="simple_white")
        return fig
    title = f"{value_name}质控图（±15% 点线为各物质自身参考带）" + (" | " + "；".join(rtxt) if rtxt else "")
    fig.update_layout(title=title, xaxis_title="实验日期",
                      yaxis_title=value_name, template="simple_white", height=470,
                      hovermode="x unified",
                      legend=dict(orientation="h", y=-0.22, x=0), margin=dict(t=60, b=80))
    return fig


def qc_area_chart(qc_df: pd.DataFrame, analytes) -> go.Figure:
    """标准品峰面积 vs 日期质控图。analytes: 单个名称或列表（多物质叠加同图）。"""
    if isinstance(analytes, str):
        analytes = [analytes]
    return _qc_overlay_chart(qc_df, list(analytes), "std_acceptor_mean", "标准品峰面积", ".3e")


def qc_width_chart(qc_df: pd.DataFrame, analytes) -> go.Figure:
    """标准品半峰宽 vs 日期质控图。"""
    if isinstance(analytes, str):
        analytes = [analytes]
    return _qc_overlay_chart(qc_df, list(analytes), "std_width_mean", "标准品半峰宽", ".4f")


def qc_rt_chart(qc_df: pd.DataFrame, analytes) -> go.Figure:
    """标准品保留时间 vs 日期质控图。"""
    if isinstance(analytes, str):
        analytes = [analytes]
    return _qc_overlay_chart(qc_df, list(analytes), "std_rt_mean", "标准品保留时间", ".4f")


# ---------- 实验环境质控（温度/湿度/柱压 各一张折线图）----------

def environment_chart(exp_df: pd.DataFrame) -> list:
    """温度/湿度/柱压分别画三张折线图（不同实验日期）。返回 [fig, fig, fig]。"""
    d = exp_df[["exp_date", "temperature", "humidity", "column_pressure"]].copy()
    for c in ["temperature", "humidity", "column_pressure"]:
        d[c] = pd.to_numeric(d[c], errors="coerce")
    d = d.sort_values("exp_date")
    figs = []
    for col, ylabel, color in [("temperature", "温度 (℃)", "#E45756"),
                               ("humidity", "湿度 (%)", "#4E79A7"),
                               ("column_pressure", "柱压", "#59A14F")]:
        sub = d[["exp_date", col]].dropna(subset=[col])
        fig = go.Figure()
        if len(sub) == 0:
            fig.update_layout(title=f"暂无{ylabel}数据（未录入）", template="simple_white")
        else:
            fig.add_trace(go.Scatter(x=sub["exp_date"], y=sub[col], mode="lines+markers",
                                     name=ylabel, line=dict(color=color, width=2),
                                     marker=dict(size=9)))
            fig.update_layout(
                title=f"实验环境质控 - {ylabel}（vs 实验日期）",
                xaxis_title="实验日期", yaxis_title=ylabel,
                template="simple_white", height=320, hovermode="x unified")
        figs.append(fig)
    return figs


def single_exp_rsd_table(qc_df: pd.DataFrame) -> pd.DataFrame:
    """
    行=实验日期，列=各物质，值=该次实验标准品RSD%。
    返回的 DataFrame 供 Streamlit 用 Styler 标红 >15%。
    """
    if len(qc_df) == 0:
        return pd.DataFrame()
    pivot = qc_df.pivot_table(index="exp_date", columns="analyte", values="std_acceptor_rsd", aggfunc="first")
    pivot = pivot.sort_index()
    return pivot


# ---------- 自定义看板 ----------

def custom_chart(metrics_df: pd.DataFrame, exp_df: pd.DataFrame,
                 x: str, y: str, color: str = None, analyte_filter: list = None) -> go.Figure:
    """
    通用自定义看板。
    metrics_df: 来自 metrics 表（含 exp_date/analyte/grp/recovery/dm/ma/A_rsd/S_rsd...）
    exp_df: 来自 experiments 表（含 temperature/humidity/pH_donor...）
    x: X轴字段（如 exp_date/temperature/humidity/pH_donor）
    y: Y轴字段（如 recovery/dm/ma）
    color: 分组颜色（如 analyte/grp）
    """
    # 合并实验级元数据
    if x in exp_df.columns and x not in metrics_df.columns:
        m = metrics_df.merge(exp_df[["exp_date", x]].drop_duplicates(), on="exp_date", how="left", suffixes=("", "_exp"))
        if f"{x}_exp" in m.columns:
            m[x] = m[f"{x}_exp"]
        metrics_df = m
    d = metrics_df.copy()
    if analyte_filter:
        d = d[d["analyte"].isin(analyte_filter)]
    if color == "analyte" or color == "grp":
        pass
    elif color:
        # 来自 exp_df 的颜色字段
        if color in exp_df.columns and color not in d.columns:
            d = d.merge(exp_df[["exp_date", color]].drop_duplicates(), on="exp_date", how="left")
    fig = px.scatter(
        d, x=x, y=y, color=color if color else None,
        hover_data=["analyte", "grp", "exp_date"],
        title=f"{y} vs {x}" + (f"（按{color}着色）" if color else ""),
        template="simple_white",
    )
    fig.update_layout(height=500)
    return fig


# ---------- 物质下钻：某物质历次回收率/转移率趋势 ----------

def analyte_trend(metrics_df: pd.DataFrame, analyte: str, metric: str = "recovery") -> go.Figure:
    """某物质历次实验的趋势图（按日期）"""
    d = metrics_df[metrics_df["analyte"] == analyte].copy()
    if len(d) == 0:
        fig = go.Figure()
        fig.update_layout(title=f"{analyte}：无数据", template="simple_white")
        return fig
    d = d.sort_values("exp_date")
    fig = px.box(d, x="exp_date", y=metric, points="all",
                 title=f"{analyte} 历次实验 {metric} 分布",
                 labels={"exp_date": "实验日期", metric: metric},
                 template="simple_white")
    fig.update_layout(height=400)
    return fig


# ---------- 单次实验 RT 散点（X=样本数，Y=RT）----------

def single_rt_scatter(rt_records_df: pd.DataFrame, exp_date: str, analyte: str = None) -> go.Figure:
    """某次实验的 RT 散点图：X=样品序号(1,2,3...)，Y=保留时间(min)，颜色=物质。
    analyte: 指定物质名则只画该物质，None 则画全部物质。
    """
    d = rt_records_df[rt_records_df["exp_date"] == exp_date].copy()
    if analyte:
        d = d[d["analyte"] == analyte]
    if len(d) == 0:
        fig = go.Figure()
        fig.update_layout(title="无RT数据", template="simple_white")
        return fig
    # 按组-相-重复排序生成样品序号
    d = d.sort_values(["grp", "phase", "replicate"]).reset_index(drop=True)
    sample_order = {}
    seq = 0
    for _, r in d.iterrows():
        key = (r["grp"], r["phase"], r["replicate"])
        if key not in sample_order:
            seq += 1
            sample_order[key] = seq
    d["样品序号"] = d.apply(lambda r: sample_order[(r["grp"], r["phase"], r["replicate"])], axis=1)
    d["样品"] = d.apply(lambda r: f"{int(r['grp'])}-{r['phase']}-{int(r['replicate'])}", axis=1)
    title = f"实验 {exp_date} 各样品保留时间" + (f" - {analyte}" if analyte else " - 全部物质")
    fig = px.scatter(d, x="样品序号", y="rt", color="analyte" if not analyte else None,
                     hover_data=["样品"],
                     title=title,
                     labels={"样品序号": "样品序号", "rt": "保留时间(min)"},
                     template="simple_white")
    fig.update_layout(height=500, xaxis=dict(dtick=1))
    return fig
