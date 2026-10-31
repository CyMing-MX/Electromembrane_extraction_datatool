# -*- coding: utf-8 -*-
"""
EME 实验数据管理系统 - 数据解析模块

职责：
1. 识别实验文件夹内的原始数据文件（优先 ALLDATA.txt，也兼容 AREA.txt + RT.txt）
2. 提取每个物质的第 1 离子对（定量离子对）作为该物质的标识
3. 正向提取样本：只保留符合 "组号-A/S-重复号" 格式的样本
4. 质控/空白：按用户输入（或默认）名称匹配所有同名行
5. 多次进样处理：样本取第一次（警告）；STD/BLK 保留全部（后续聚合）
6. N/A 视为 0
7. 输出干净的长表 DataFrame，可直接进入计算

输出 DataFrame 列：
    sample       样本名（原始，如 1-A-1 / Accept-STD-100ppb-1）
    group        组号（int，仅样本有；STD/BLK 为 NaN）
    phase        相（A=受体/S=供体，仅样本有；STD/BLK 为 NaN）
    replicate    重复号（int，仅样本有；STD/BLK 为 NaN）
    category     类别：sample / std_acceptor / std_donor / blk_acceptor / blk_donor
    analyte      物质名（=组名，如 Methotrexate）
    component    离子对全名（如 Methotrexate-1）
    area         峰面积（N/A→0）
    rt           保留时间（N/A→0；若无RT文件则为NaN）
    height       峰高（仅ALLDATA有，否则NaN）
    width50      半峰宽（仅ALLDATA有，否则NaN）
    injection_idx 该样本该离子对的第几次进样（1起）
"""
from __future__ import annotations
import re
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional

import pandas as pd

try:
    from . import config
except ImportError:  # 直接运行本文件时
    import config


# ---------- 文件识别 ----------

@dataclass
class RawFiles:
    """一次实验找到的原始文件"""
    folder: Path
    alldata: Optional[Path] = None   # *-ALLDATA.txt（长表，优先）
    area: Optional[Path] = None      # *-AREA.txt（宽表）
    rt: Optional[Path] = None        # *-RT.txt（宽表）
    # 文件前缀日期（如 20251112），从文件名提取
    file_date: str = ""

    @property
    def has_data(self) -> bool:
        return self.alldata is not None or (self.area is not None and self.rt is not None)


def find_raw_files(folder: str | Path) -> RawFiles:
    """扫描实验文件夹，识别原始数据文件"""
    folder = Path(folder)
    rf = RawFiles(folder=folder)
    if not folder.is_dir():
        return rf
    for p in sorted(folder.glob("*.txt")):
        name = p.name.upper()
        if "ALLDATA" in name:
            rf.alldata = p
        elif name.endswith("-AREA.TXT") or "AREA" in name:
            rf.area = p
        elif name.endswith("-RT.TXT") or re.search(r"[\-_]RT[\-_\.]", name):
            rf.rt = p
        # 从文件名提取日期前缀（形如 20251112-ALLDATA.txt）
        m = re.match(r"^(\d{8})", p.name)
        if m and not rf.file_date:
            rf.file_date = m.group(1)
    return rf


# ---------- ALLDATA（长表）解析 ----------

# ALLDATA 表头列名（容忍空格/大小写）
_ALLDATA_COLS = {
    "sample":   ["sample name", "sample"],
    "comp":     ["component name", "component"],
    "group":    ["component group name", "group name", "group"],
    "area":     ["area"],
    "height":   ["height"],
    "rt":       ["retention time", "rt", "retentiontime"],
    "width50":  ["width at 50%", "width50", "width_50", "fwhm"],
    "sn":       ["signal / noise", "signal /noise", "signal/ noise", "signal/noise",
                 "signal noise", "signal to noise", "s / n", "s/n", "sn", "信噪比"],
}


def _match_column(actual: str, candidates: list[str]) -> bool:
    a = actual.strip().lower()
    return a in [c.lower() for c in candidates] or any(c.lower() in a for c in candidates)


def _col_map(columns: list[str]) -> dict:
    """把实际表头映射到标准字段名"""
    mapping = {}
    used = set()
    for std_field, cands in _ALLDATA_COLS.items():
        for col in columns:
            if col in used:
                continue
            if _match_column(col, cands):
                mapping[std_field] = col
                used.add(col)
                break
    return mapping


def _read_csv_flexible(path, sep="\t"):
    """读制表符导出，编码自适应：utf-8(-sig) → gbk → utf-16（色谱工作站常见多编码导出）"""
    for enc in ("utf-8-sig", "utf-8", "gbk", "utf-16"):
        try:
            return pd.read_csv(path, sep=sep, dtype=str, encoding=enc)
        except (UnicodeDecodeError, UnicodeError):
            continue
    return pd.read_csv(path, sep=sep, dtype=str, encoding="utf-8", encoding_errors="replace")


def _to_float(v) -> float:
    """N/A / 空 / NaN → 0（按需求：N/A 视为 0）"""
    if v is None:
        return 0.0
    if isinstance(v, float):
        # pandas 读到的 NaN（float nan）
        import math
        return 0.0 if math.isnan(v) else v
    if isinstance(v, int):
        return float(v)
    s = str(v).strip()
    if s == "" or s.upper() in ("N/A", "NA", "NAN", "NONE", "-"):
        return 0.0
    s = s.replace(",", "")  # 千分位（"1,234.5"）：先剥再转，否则静默归 0
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def _normalize_analyte_name(name: str) -> str:
    """规范化物质名：与信息库/图片名一致。
    Component Group Name 常用点（5.Fluorouracil），而信息库/图片用连字符（5-Fluorouracil）。
    将物质名中的点替换为连字符，统一格式。
    """
    if name is None:
        return ""
    s = str(name).strip()
    # 5.Fluorouracil → 5-Fluorouracil；N.norimatinib → N-norimatinib
    s = s.replace(".", "-")
    return s


def parse_alldata(path: str | Path) -> pd.DataFrame:
    """解析 ALLDATA 长表，返回规范化的中间表（未做样本筛选）。
    列：sample, analyte, component, area, rt, height, width50
    物质名优先从 Component Name 去尾部离子对号提取（避免 Group Name 的点号不一致）。
    """
    df = _read_csv_flexible(path)
    df.columns = [str(c).strip() for c in df.columns]
    cm = _col_map(list(df.columns))
    if "sample" not in cm:
        raise ValueError(f"ALLDATA 文件未找到 Sample Name 列：{path}")
    comp_col = cm.get("comp")
    out = pd.DataFrame()
    out["sample"] = df[cm["sample"]].astype(str).str.strip()
    out["component"] = df[comp_col].astype(str).str.strip() if comp_col else ""
    # 物质名：优先从 Component Name 去尾部 -数字 提取（与信息库一致），并规范化点→连字符
    if comp_col:
        out["analyte"] = out["component"].str.replace(r"-\d+$", "", regex=True).map(_normalize_analyte_name)
    else:
        group_col = cm.get("group")
        out["analyte"] = df[group_col].astype(str).str.strip().map(_normalize_analyte_name) if group_col else ""
    out["area"] = df[cm["area"]].map(_to_float) if "area" in cm else 0.0
    out["rt"] = df[cm["rt"]].map(_to_float) if "rt" in cm else float("nan")
    out["height"] = df[cm["height"]].map(_to_float) if "height" in cm else float("nan")
    out["width50"] = df[cm["width50"]].map(_to_float) if "width50" in cm else float("nan")
    out["sn"] = df[cm["sn"]].map(_to_float) if "sn" in cm else float("nan")
    return out


# ---------- AREA + RT（宽表）解析 ----------

def parse_wide_area_rt(area_path: str | Path, rt_path: str | Path | None,
                       ion_pair: int = config.DEFAULT_ION_PAIR) -> pd.DataFrame:
    """解析 AREA（+可选RT）宽表，返回与 parse_alldata 同结构的中间表。
    宽表列：Sample Name + 各 '物质-离子对' 列。
    物质名从列名提取（去尾部 -数字）。仅取每个物质的第 1 离子对。
    """
    a = _read_csv_flexible(area_path)
    a.columns = [str(c).strip() for c in a.columns]
    sample_col = a.columns[0]
    # 收集列对应的 (analyte, component, 离子对号)
    col_info = []  # [(col_name, analyte, component, ion_idx)]
    ion_seen = {}  # analyte -> set of ion_idx already kept
    for col in a.columns[1:]:
        m = re.match(r"^(.+?)-(\d+)$", col)
        if not m:
            continue
        analyte, ion = m.group(1), int(m.group(2))
        # 仅保留每个物质的第 1 离子对
        if ion != ion_pair:
            continue
        analyte = _normalize_analyte_name(analyte)  # 规范化点→连字符
        col_info.append((col, analyte, col, ion))
    # 构建长表
    records = []
    for _, row in a.iterrows():
        sample = str(row[sample_col]).strip()
        for col, analyte, comp, ion in col_info:
            rec = {
                "sample": sample, "analyte": analyte, "component": comp,
                "area": _to_float(row[col]),
                "rt": float("nan"), "height": float("nan"), "width50": float("nan"),
                "sn": float("nan"),
            }
            records.append(rec)
    out = pd.DataFrame(records)
    # 合并 RT
    if rt_path is not None and Path(rt_path).exists():
        r = _read_csv_flexible(rt_path)
        r.columns = [str(c).strip() for c in r.columns]
        rsample = r.columns[0]
        for col, analyte, comp, ion in col_info:
            if col in r.columns:
                # 按 sample+col 建 map
                rt_map = dict(zip(r[rsample].astype(str).str.strip(), r[col]))
                mask = out["component"] == col
                out.loc[mask, "rt"] = out.loc[mask, "sample"].map(lambda s: _to_float(rt_map.get(s, 0.0)))
    return out


# ---------- 样本名解析（组号-相-重复）----------

_PATTERN_COMPILE_CACHE = {}


# 旧版预设正则（相位仅 [AS]）→ 新版（含 D）：存库的全局/实验级格式透明升级，
# 无需用户重新设置；D 命名随后由解析归一化为 S
_LEGACY_PHASE_PATTERNS = {
    r"^(?P<g>\d+)-(?P<ph>[AS])-(?P<r>\d+)$": r"^(?P<g>\d+)-(?P<ph>[ASD])-(?P<r>\d+)$",
    r"^(?P<ph>[AS])(?P<g>\d+)-(?P<r>\d+)$": r"^(?P<ph>[ASD])(?P<g>\d+)-(?P<r>\d+)$",
}


def compile_sample_pattern(fmt: str):
    """编译样品名称正则（带缓存）。必须含命名组 g/ph/r，否则回退内置默认。"""
    if fmt in _LEGACY_PHASE_PATTERNS:
        fmt = _LEGACY_PHASE_PATTERNS[fmt]
    rx = _PATTERN_COMPILE_CACHE.get(fmt)
    if rx is None:
        try:
            rx = re.compile(fmt)
            if not {"g", "ph", "r"}.issubset(rx.groupindex.keys()):
                rx = re.compile(config.SAMPLE_NAME_PATTERN)
        except re.error:
            rx = re.compile(config.SAMPLE_NAME_PATTERN)
        _PATTERN_COMPILE_CACHE[fmt] = rx
    return rx


def get_sample_re():
    """当前全局设置的样品名称正则（设置文件按 mtime 缓存，改动自动生效）"""
    try:
        from . import settings as _S
    except ImportError:
        import settings as _S
    return compile_sample_pattern(_S.get_sample_format())


def parse_sample_name(name: str, pattern: str | None = None):
    """返回 (group, phase, replicate) 或 None（不符合样本格式）。
    pattern 缺省用全局设置的正则；传入 pattern 时用指定格式（如实验保存的覆盖格式）。"""
    rx = compile_sample_pattern(pattern) if pattern else get_sample_re()
    m = rx.match(name or "")
    if not m:
        return None
    try:
        g, ph, r = m.group("g"), m.group("ph"), m.group("r")
    except Exception:  # 兼容缺命名组的模式（按位置组）
        try:
            g, ph, r = m.group(1), m.group(2), m.group(3)
        except Exception:
            return None
    try:
        _ph = str(ph).upper()
        if _ph == "D":  # 供体相 D 命名 → 系统内部统一记作 S
            _ph = "S"
        return int(g), _ph, int(r)
    except (TypeError, ValueError):
        return None


# ---------- 主入口：解析一次实验 ----------

@dataclass
class ParseResult:
    """一次实验的解析结果"""
    df: pd.DataFrame               # 干净长表（含 sample/std/blk 各行）
    analytes: list[str]            # 本实验涉及的所有物质名（有序）
    groups: list[int]              # 本实验的样品组号列表
    n_replicates: int              # 最大重复数（通常3）
    warnings: list[str] = field(default_factory=list)
    # 实际用于匹配的 {std_acceptor, std_donor, blk_acceptor, blk_donor}
    qc_names: dict = field(default_factory=dict)
    file_date: str = ""
    # 实际生效的样品名格式正则（含格式自愈后的结果）；下游（rt_records/Excel 有效数/排序）
    # 解析样品名时必须用它，不得退回全局设置——否则实验级格式覆盖丢失
    sample_format: str = ""
    # 原始文件中出现的全部样品名（含洗针/未识别等，按出现顺序）——供质控名下拉框完整选择
    all_sample_names: list = field(default_factory=list)


def parse_experiment(
    folder: str | Path,
    qc_names: Optional[dict] = None,
    use_sample: bool = True,
    ion_pair: int = config.DEFAULT_ION_PAIR,
    sample_pattern: Optional[str] = None,
) -> ParseResult:
    """
    解析一次实验，输出干净长表。

    参数
    ----
    folder : 实验文件夹
    qc_names : dict，键 std_acceptor/std_donor/blk_acceptor/blk_donor；
               缺省用 config 默认名；值为 None 表示该项不使用（例如两相溶液相同，供体标准品=受体标准品）
    use_sample : 是否纳入供体相样品（影响是否计算转移率；解析阶段只标记，不删数据）
    ion_pair : 定量离子对号，默认 1

    返回
    ----
    ParseResult
    """
    folder = Path(folder)
    rf = find_raw_files(folder)
    if not rf.has_data:
        raise FileNotFoundError(f"实验文件夹未找到可用数据文件：{folder}")

    warnings = []
    qc_names = qc_names or {}

    # —— 1. 读原始长表 ——
    if rf.alldata:
        raw = parse_alldata(rf.alldata)
        source = "ALLDATA"
    else:
        raw = parse_wide_area_rt(rf.area, rf.rt, ion_pair=ion_pair)
        source = "AREA+RT"

    # —— 2. 仅保留第 N 离子对（ALLDATA 含全部离子对；AREA 宽表已过滤）——
    if source == "ALLDATA":
        def _ion_idx(comp: str) -> int:
            m = re.search(r"-(\d+)$", comp or "")
            return int(m.group(1)) if m else 0
        raw["_ion"] = raw["component"].map(_ion_idx)
        before = len(raw)
        # 同一物质若无指定离子对，保留 ion 最大者退化处理：仅保留 == ion_pair
        raw = raw[raw["_ion"] == ion_pair].copy()
        dropped = before - len(raw)
        if dropped:
            warnings.append(f"已保留第 {ion_pair} 离子对，过滤掉 {dropped} 行其它离子对数据")
        raw.drop(columns=["_ion"], inplace=True)

    # —— 3. 分类：sample / std_acceptor / std_donor / blk_acceptor / blk_donor ——
    _sample_rx = compile_sample_pattern(sample_pattern) if sample_pattern else get_sample_re()

    # —— 格式自愈：用户未显式指定格式、且当前格式识别不出任何样品名时，
    #    依次尝试内置预设（如 A1-1），命中则改用并在警告中说明（可在设置页设为默认）——
    if sample_pattern is None:
        _hit_global = any(_sample_rx.match(n) for n in raw["sample"].unique())
        if not _hit_global:
            try:
                from . import settings as _S
            except ImportError:
                import settings as _S
            for _lbl_p, _rx_p in _S.SAMPLE_FORMAT_PRESETS.items():
                if _rx_p == _sample_rx.pattern:
                    continue
                if any(compile_sample_pattern(_rx_p).match(n) for n in raw["sample"].unique()):
                    _sample_rx = compile_sample_pattern(_rx_p)
                    warnings.append(
                        f"样品名不符合当前默认格式，已自动按「{_lbl_p}」解析"
                        "（建议到 ⚙️设置 → 样品名称格式 设为默认，或在新增实验页确认）")
                    break

    def classify(name: str) -> str:
        if _sample_rx.match(name or "") is not None:
            return "sample"
        return "other"

    raw["category"] = raw["sample"].map(classify)

    # 质控/空白匹配（宽松自动识别：名称含 STD → 标准品候选；含 BLK/BLANK → 空白候选）
    # 候选名按进样顺序（非字母排序）：序列中先跑的质控样更可能是本实验主标准品/空白
    _name_order = list(dict.fromkeys(raw["sample"].tolist()))
    all_names = set(_name_order)
    std_names = [n for n in _name_order
                 if "std" in n.lower() and _sample_rx.match(n) is None]
    blk_names = [n for n in _name_order
                 if ("blk" in n.lower() or "blank" in n.lower())
                 and _sample_rx.match(n) is None]
    if not std_names and not blk_names:
        # 拼音缩写习惯（bz=标准品, kb=空白）作为兜底识别（仅在无 STD/BLK 候选时）
        std_names = [n for n in _name_order
                     if _sample_rx.match(n) is None
                     and (n.lower().startswith("bz") or "标准" in n)]
        blk_names = [n for n in _name_order
                     if _sample_rx.match(n) is None
                     and (n.lower().startswith("kb") or "空白" in n)]
        if std_names or blk_names:
            warnings.append("按拼音缩写识别质控候选（bz=标准品 / kb=空白），请在界面确认")

    def _auto_fill(role: str, configured: str) -> str:
        """配置名不在数据中时，用自动识别的 STD/BLK 名回填"""
        if configured in all_names:
            return configured
        if role.startswith("std") and std_names:
            pick = std_names[0] if len(std_names) == 1 else std_names[0]
            warnings.append(f"自动识别{role}='{pick}'（配置名'{configured}'不在数据中），请在界面确认")
            return pick
        if role.startswith("blk") and blk_names:
            pick = blk_names[0]
            warnings.append(f"自动识别{role}='{pick}'（配置名'{configured}'不在数据中），请在界面确认")
            return pick
        return configured

    qc_keys = ["std_acceptor", "std_donor", "blk_acceptor", "blk_donor"]
    qc_used = {}
    for k in qc_keys:
        nm = qc_names.get(k)
        if isinstance(nm, (list, tuple)):
            # 多标准品：保留数据中存在的名字；全缺失则回退自动识别单个
            present = [n for n in nm if n in all_names]
            if present:
                qc_used[k] = present
                continue
            pick = (std_names or blk_names or [None])[0]
            if pick:
                warnings.append(f"多标准品组合的名字均不在数据中，{k} 自动识别为 '{pick}'，请在界面确认")
                qc_used[k] = pick
            else:
                qc_used[k] = list(nm)
            continue
        if not nm:
            # 用默认
            nm = {
                "std_acceptor": config.DEFAULT_STD_ACCEPTOR,
                "std_donor": config.DEFAULT_STD_DONOR,
                "blk_acceptor": config.DEFAULT_BLK_ACCEPTOR,
                "blk_donor": config.DEFAULT_BLK_DONOR,
            }[k]
        qc_used[k] = _auto_fill(k, nm)
    if len(std_names) >= 2:
        warnings.append(f"检测到多个标准品名 {std_names}，已默认使用第一个，"
                        "如有供体相/受体相两种标准品请在界面分别指定")
    if std_names or blk_names:
        warnings.append(f"自动识别候选 — 标准品: {std_names or '无'}；空白: {blk_names or '无'}")
    _other_names = [n for n in _name_order
                    if _sample_rx.match(n) is None and n not in std_names and n not in blk_names]
    if _other_names:
        warnings.append(f"以下非样品名未被识别为质控（可在质控下拉框手动指定）: {_other_names[:8]}")

    # 一个样本名可能同时被指定为多个角色（两相溶液相同时，供体=受体同名）
    # 处理：若 std_donor == std_acceptor，则该名样本同时归两类（复制行）；
    # 为避免重复，采用"主角色优先"：按 std_acceptor > std_donor > blk_acceptor > blk_donor
    # 但两相相同时用户填同名是合理的，我们需要保留双轨计算所需的数据。
    # 策略：分类时，同一名字命中多个角色 → 分别为每个角色保留一份行（category 不同）。
    # —— 同名双计防护：质控名若同时匹配样品格式，该行会既进样品又进质控（回收率分子分母
    #    被同一数据污染）——明确警告，提示用户核对质控下拉选择
    _qc_names_flat = {n for k in qc_keys for n in
                      (list(qc_used[k]) if isinstance(qc_used.get(k), (list, tuple)) else ([qc_used[k]] if qc_used.get(k) else []))}
    _dup_qc = sorted(n for n in _qc_names_flat if _sample_rx.match(n or "") is not None)
    if _dup_qc:
        warnings.append(f"⚠️ 质控名 {'、'.join(_dup_qc[:5])} 同时符合样品名格式：这些进样既被计入样品组又被计入质控，"
                        "请核对质控下拉框是否选错（正常标准品/空白名不含 组-相-重复 结构）")
    role_rows = []
    # 先处理 sample（已标记）
    role_rows.append(raw[raw["category"] == "sample"].copy())
    for k in qc_keys:
        nm = qc_used[k]
        # 支持列表：多标准品组合时一个类别保留多份标准品的行（sample 列保留原名）
        names_l = list(nm) if isinstance(nm, (list, tuple)) else ([nm] if nm else [])
        for _n in names_l:
            sub = raw[raw["sample"] == _n].copy()
            if len(sub) == 0:
                warnings.append(f"未在数据中找到 {k}='{_n}' 的样本行")
                continue
            sub["category"] = k
            role_rows.append(sub)
    df = pd.concat(role_rows, ignore_index=True) if role_rows else pd.DataFrame()

    # —— 4. 多次进样处理 ——
    # sample：同一 (sample, component) 多行 → 取第一次（injection_idx 最小），并警告
    # std/blk：保留全部，后续 calculator 聚合时取平均
    if len(df) > 0:
        df = df.sort_values(["sample", "component"], kind="stable").reset_index(drop=True)
        # 给每个 (sample, component) 编 injection_idx
        # 按类别分别编针号：供体相/受体相同名标准品时，各自的 #N 对应同一物理进样，
        # 剔除任一相的 #N 即可同时命中两份拷贝（同名标准品自动同时应用于两相）
        df["injection_idx"] = df.groupby(["category", "sample", "component"]).cumcount() + 1
        # 对 sample 类，仅保留 injection_idx==1
        is_sample = df["category"] == "sample"
        dup_mask = is_sample & (df["injection_idx"] > 1)
        if dup_mask.any():
            # 按样本统计重复
            dups = df[dup_mask].groupby(["sample", "component"]).size()
            for (s, c), n in dups.items():
                warnings.append(f"样本 {s} 的 {c} 出现 {n+1} 次进样，已取第 1 次")
        df = df[~dup_mask].copy()

    # —— 5. 拆分样本的 group/phase/replicate（向量化提取，避免逐行 apply）——
    if len(df) > 0:
        is_smp = df["category"] == "sample"
        ext = df["sample"].where(is_smp, "").str.extract(_sample_rx.pattern)
        ext = ext.rename(columns={"g": "group", "ph": "phase", "r": "replicate"})
        df["group"] = pd.to_numeric(ext["group"], errors="coerce").astype("Int64")
        # 供体相 D 命名 → 内部统一记作 S（下游 相位判断/M分表/剔除 全按 A/S 消费）
        df["phase"] = ext["phase"].str.upper().replace({"D": "S"})
        df["replicate"] = pd.to_numeric(ext["replicate"], errors="coerce").astype("Int64")
    else:
        df["group"] = pd.Series(dtype="object")
        df["phase"] = pd.Series(dtype="object")
        df["replicate"] = pd.Series(dtype="object")

    # —— 6. 汇总信息 ——
    analytes = sorted(df["analyte"].dropna().unique().tolist()) if len(df) else []
    sample_df = df[df["category"] == "sample"] if len(df) else df
    groups = sorted(sample_df["group"].dropna().astype(int).unique().tolist()) if len(sample_df) else []
    n_rep = int(sample_df["replicate"].max()) if len(sample_df) and sample_df["replicate"].notna().any() else 0

    return ParseResult(
        df=df.reset_index(drop=True),
        analytes=analytes,
        groups=groups,
        n_replicates=n_rep,
        warnings=warnings,
        qc_names={k: (v[0] if isinstance(v, (list, tuple)) and v else v) for k, v in qc_used.items()},
        file_date=rf.file_date,
        sample_format=getattr(_sample_rx, "pattern", "") or "",
        all_sample_names=list(dict.fromkeys(raw["sample"].tolist())),
    )


# ---------- 命令行自检 ----------
if __name__ == "__main__":
    import sys
    fld = sys.argv[1] if len(sys.argv) > 1 else "../data/experiments/20251112-8M"
    res = parse_experiment(fld)
    print("=" * 60)
    print(f"实验文件夹: {fld}")
    print(f"识别物质数: {len(res.analytes)}")
    print(f"样品组号: {res.groups}")
    print(f"最大重复数: {res.n_replicates}")
    print(f"质控/空白名: {res.qc_names}")
    print(f"数据总行数: {len(res.df)}")
    print(f"各类别计数:\n{res.df['category'].value_counts()}")
    print(f"\n警告({len(res.warnings)}):")
    for w in res.warnings[:10]:
        print("  -", w)
    print("\n样本行预览（前10）:")
    print(res.df[res.df['category'] == 'sample'].head(10).to_string())
