# -*- coding: utf-8 -*-
"""
EME 实验数据管理系统 - 电流记录读取与绘图

电流文件存放：data/current/*.xlsx
- 每个 xlsx 可含多个 sheet
- sheet 命名规则：
  - 非测试萃取时间/多因素：sheet 名 = 实验文件夹名（如 20251112-SM-8M）
  - 测试萃取时间/多因素：sheet 名 = 文件夹名 + 萃取时间（如 20251112-SM-5M-25min）
- 每个 sheet 格式统一：A列=序号, 样本列=电流值+单位（如 "0.162 mA DC"）
- 时间 x 轴 = 序号 / 120（单位：分钟）
- 电流单位从样本列文本解析（mA 或 uA）

遍历 current 文件夹下所有 xlsx 的所有 sheet 建立索引，按 sheet 名查找。
"""
from __future__ import annotations
import re
from pathlib import Path
from typing import Optional

import streamlit as st
import pandas as pd
import plotly.graph_objects as go

try:
    from . import config
except ImportError:
    import config

CURRENT_DIR = config.DATA_DIR / "current"


def _parse_current_value(text) -> tuple[float, str]:
    """从 '0.162 mA DC' 解析出 (数值, 单位)。失败返回 (None, '')。
    支持科学计数法（1.62E-4 mA）；µ/μ 两种微符号均识别。"""
    if text is None:
        return None, ""
    s = str(text).strip()
    m = re.search(r"(-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)\s*(mA|uA|µA|μA|A)", s, re.IGNORECASE)
    if not m:
        return None, ""
    val = float(m.group(1))
    unit = m.group(2)
    # 统一单位写法
    if unit.lower() in ("ua", "µa", "μa"):
        unit = "uA"
    elif unit.lower() == "ma":
        unit = "mA"
    elif unit.lower() == "a":
        unit = "A"
    return val, unit


def _dir_signature() -> tuple:
    """current 目录签名（文件名+mtime 列表），用作缓存 key：增删/改文件自动失效"""
    if not CURRENT_DIR.exists():
        return ()
    sig = []
    for xlsx in sorted(CURRENT_DIR.glob("*.xlsx")):
        try:
            if xlsx.name.startswith("~$"):
                continue
            sig.append((xlsx.name, xlsx.stat().st_mtime))
        except OSError:
            continue
    return tuple(sig)


@st.cache_data(show_spinner=False)
def _list_all_sheets_cached(sig: tuple) -> dict:
    index = {}
    for name, _mt in sig:
        xlsx = CURRENT_DIR / name
        try:
            xl = pd.ExcelFile(xlsx)
        except Exception:
            continue
        for sn in xl.sheet_names:
            index[sn] = xlsx
    return index


def list_all_sheets() -> dict:
    """
    遍历 current 文件夹所有 xlsx，返回 {sheet名: 文件路径} 索引。
    （按目录签名缓存，同一会话内反复调用不再逐个打开 xlsx）
    """
    return dict(_list_all_sheets_cached(_dir_signature()))


@st.cache_data(show_spinner=False)
def _list_all_sheets_detailed_cached(sig: tuple) -> list:
    out = []
    for name, _mt in sig:
        xlsx = CURRENT_DIR / name
        try:
            xl = pd.ExcelFile(xlsx)
        except Exception:
            continue
        for sn in xl.sheet_names:
            out.append({"file": name, "sheet": sn})
    return out


def list_all_sheets_detailed() -> list:
    """
    遍历 current 文件夹所有 xlsx，返回 [{file, sheet}] 详细清单（供用户选择）。
    """
    return list(_list_all_sheets_detailed_cached(_dir_signature()))


def _parse_sheet_df(df: pd.DataFrame, source_name: str) -> Optional[dict]:
    """解析单个 sheet 的电流数据（核心解析逻辑）"""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    # A列=序号, 样本列=电流值
    a_col = None
    sample_col = None
    for c in df.columns:
        if str(c).strip().upper() == "A":
            a_col = c
        if str(c).strip() in ("样本", "Sample", "样本 "):
            sample_col = c
    if a_col is None:
        a_col = df.columns[0]
    if sample_col is None:
        sample_col = df.columns[1] if len(df.columns) > 1 else None
    if sample_col is None:
        return None

    seqs = []
    raw = []  # (数值, 单位) 原始采集，循环后统一处理单位
    for _, row in df.iterrows():
        seq = row[a_col]
        try:
            seq = float(seq)
        except (ValueError, TypeError):
            continue
        val, unit = _parse_current_value(row[sample_col])
        if val is None:
            continue
        seqs.append(seq)
        raw.append((val, unit))
    if not seqs:
        return None
    time_min = [s / 120.0 for s in seqs]
    units = {u for _, u in raw}
    if len(units) > 1:
        # 同一 sheet 混用单位（如既有 mA 又有 uA）→ 统一换算成 mA，否则原始数值混画差 1000 倍
        currents = [v / 1000.0 if u == "uA" else (v * 1000.0 if u == "A" else v) for v, u in raw]
        unit = "mA"
    else:
        currents = [v for v, _ in raw]
        unit = next(iter(units)) if units else "mA"
    return {
        "time_min": time_min,
        "current": currents,
        "unit": unit,
        "source_file": source_name,
        "n_points": len(seqs),
    }


@st.cache_data(show_spinner=False)
def _load_current_pair_cached(file_name: str, sheet_name: str, mtime: float) -> Optional[dict]:
    path = CURRENT_DIR / file_name
    try:
        df = pd.read_excel(path, sheet_name=sheet_name)
    except Exception:
        return None
    return _parse_sheet_df(df, file_name)


def load_current_pair(file_name: str, sheet_name: str) -> Optional[dict]:
    """
    按 (文件名, sheet名) 精确加载电流数据（用户在新增实验时选择的映射）。
    （mtime 入缓存 key：文件更新自动失效）
    """
    path = CURRENT_DIR / file_name
    if not path.exists():
        return None
    try:
        mtime = path.stat().st_mtime
    except OSError:
        return None
    return _load_current_pair_cached(file_name, sheet_name, mtime)


def load_current_sheet(sheet_name: str) -> Optional[dict]:
    """
    按 sheet 名加载电流数据（兼容按命名规则自动匹配）。
    返回 {"time_min": [...], "current": [...], "unit": "mA", "source_file": "xxx.xlsx"}
    若未找到返回 None。
    """
    index = list_all_sheets()
    if sheet_name not in index:
        # 模糊匹配（sheet名可能是文件夹名+时间，尝试包含匹配）
        candidates = [s for s in index if sheet_name in s or s in sheet_name]
        if candidates:
            sheet_name = candidates[0]
        else:
            return None
    xlsx = index[sheet_name]
    try:
        df = pd.read_excel(xlsx, sheet_name=sheet_name)
    except Exception:
        return None
    return _parse_sheet_df(df, xlsx.name)


def find_current_for_experiment(folder_name: str, exp_type: str = None,
                                extract_times: list = None) -> dict:
    """
    查找某次实验的电流数据。
    - 非测试时间/多因素：sheet 名 = folder_name（一次实验一条）
    - 测试时间/多因素：sheet 名 = folder_name + 萃取时间（每组一条）

    返回 {sheet_name: current_data_dict}（可能多条，测试时间时每组一条）
    """
    index = list_all_sheets()
    result = {}
    is_time_exp = exp_type in ("time", "all")
    if is_time_exp and extract_times:
        # 每个萃取时间一个 sheet
        for t in extract_times:
            t_str = str(int(t)) if float(t) == int(t) else str(t)
            sn = f"{folder_name}-{t_str}min"
            # 也尝试不带 min
            for candidate in [sn, f"{folder_name}-{t_str}"]:
                if candidate in index:
                    data = load_current_sheet(candidate)
                    if data:
                        result[candidate] = data
                        break
            else:
                # 模糊
                data = load_current_sheet(sn)
                if data:
                    result[sn] = data
    else:
        # sheet 名 = folder_name
        if folder_name in index:
            data = load_current_sheet(folder_name)
            if data:
                result[folder_name] = data
        else:
            # 模糊匹配：子串
            matched = False
            data = load_current_sheet(folder_name)
            if data:
                for s in index:
                    if folder_name in s or s in folder_name:
                        d = load_current_sheet(s)
                        if d:
                            result[s] = d
                            matched = True
                            break
            # 日期前缀匹配：文件夹名和 sheet 名都以 8 位日期开头（如 20261030）
            if not matched:
                date_prefix = folder_name[:8] if len(folder_name) >= 8 else folder_name
                _cands = [s for s in index if s.startswith(date_prefix)]
                if len(_cands) == 1:
                    d = load_current_sheet(_cands[0])
                    if d:
                        result[_cands[0]] = d
                        matched = True
                elif len(_cands) > 1:
                    # 多个同名日期前缀 sheet：全选（可能是同一天多组实验）
                    for s in _cands:
                        d = load_current_sheet(s)
                        if d:
                            result[s] = d
                            matched = True
    return result


def plot_current(data: dict, title: str = "") -> go.Figure:
    """绘制电流-时间曲线"""
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=data["time_min"], y=data["current"],
        mode="lines", line=dict(color="#4472C4", width=1),
        name=f"电流({data['unit']})"))
    # 标注最大/最小/平均
    cur = data["current"]
    if cur:
        max_v, min_v, mean_v = max(cur), min(cur), sum(cur) / len(cur)
        fig.add_hline(y=mean_v, line_dash="dot", line_color="gray",
                      annotation_text=f"平均={mean_v:.3f}{data['unit']}")
    fig.update_layout(
        title=title or f"电流-时间曲线（来源：{data['source_file']}）",
        xaxis_title="时间 (min)", yaxis_title=f"电流 ({data['unit']})",
        template="simple_white", height=450)
    return fig


if __name__ == "__main__":
    print("===== current 文件夹索引 =====")
    idx = list_all_sheets()
    for sn, path in idx.items():
        print(f"  {sn}  ← {path.name}")
    if idx:
        sn0 = list(idx.keys())[0]
        data = load_current_sheet(sn0)
        if data:
            print(f"\n===== {sn0} =====")
            print(f"  数据点: {data['n_points']}, 单位: {data['unit']}, 来源: {data['source_file']}")
            print(f"  时间范围: {data['time_min'][0]:.3f} ~ {data['time_min'][-1]:.3f} min")
            print(f"  电流范围: {min(data['current']):.4f} ~ {max(data['current']):.4f} {data['unit']}")
