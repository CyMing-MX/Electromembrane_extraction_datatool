# -*- coding: utf-8 -*-
"""
EME 实验数据管理系统 - 信息库读取（analyte.xlsx / SLM.xlsx）

analyte.xlsx 结构（已确认）：
  - chemical 工作表：物质性质宽表，主键 Englishname（备 name）。列含 Logp/HBA/HBD/分子量 等
  - isoelectric 工作表：name | pH | 0.0 | 0.1 | ... | 14.0，第2列=Charge
  - LogD 工作表：同结构，第2列=logD

SLM.xlsx 结构（已确认）：
  - LF-液膜类 / DES-深共晶类：主SLM（主键 试剂名称）
  - IL-离子液体类 / O-其他：添加剂（主键 试剂名称）

设计：
  - 列类型自适应：每列尝试转 float，能转=连续型（可排序/散点），不能=分类型
  - LogD/Charge 宽表：按 pH 取最近列
  - 物质名匹配：analyte 用 Englishname 与 data 中 analyte（=组名）对应；遇 . 与 - 不一致做归一化
"""
from __future__ import annotations
import math
from functools import lru_cache
from pathlib import Path
from typing import Optional

import pandas as pd

try:
    from . import config
    from . import settings
except ImportError:
    import config
    import settings


# ---------- 工具：N/A / null / 空字符串 → NaN ----------

def _clean_val(v):
    if v is None:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    s = str(v).strip()
    if s == "" or s.lower() in ("n/a", "na", "nan", "none", "null", "-"):
        return None
    return s


def _try_float(v):
    """尝试转 float，失败返回 None"""
    s = _clean_val(v)
    if s is None:
        return None
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def _is_continuous_column(series: pd.Series) -> bool:
    """判断一列是否为连续型（≥60% 可转 float）"""
    total = len(series)
    if total == 0:
        return False
    n_float = sum(_try_float(v) is not None for v in series)
    return n_float / total >= 0.6


# ============ analyte 信息库 ============

@lru_cache(maxsize=1)
def _load_analyte_chemical() -> Optional[pd.DataFrame]:
    path = config.ANALYTE_FILE
    if not path.exists():
        return None
    try:
        df = pd.read_excel(path, sheet_name="chemical")
    except Exception:
        # 兼容不同命名的性质表（"drug" 等）：找含名称主键列的工作表，退化读第一个
        try:
            xl = pd.ExcelFile(path)
            _name_keys = {"englishname", "english_name", "input-englishname", "input_englishname",
                          "name", "物质", "名称"}
            _sn = xl.sheet_names[0]
            for _s in xl.sheet_names:
                _cols_l = {str(c).strip().lower() for c in xl.parse(_s, nrows=5).columns}
                if _cols_l & _name_keys:
                    _sn = _s
                    break
            df = pd.read_excel(path, sheet_name=_sn)
        except Exception:
            return None
    df.columns = [str(c).strip() for c in df.columns]
    # 清洗：null/N/A → None
    for c in df.columns:
        df[c] = df[c].map(_clean_val)
    # 识别主键列（优先 Englishname，其次 name；兼容 "Input-englishname" 等变体）
    _cols_l = {str(c).strip().lower(): c for c in df.columns}
    for cand in ["englishname", "english_name", "input-englishname", "input_englishname",
                 "name", "物质", "名称"]:
        if cand in _cols_l:
            df = df.set_index(_cols_l[cand])
            break
    return df


# 旧列名 → 新列名（analyte.xlsx 2026-08 改版：chemical→drug，列名加 C- 前缀）
# 用于详情展示列的兼容映射；新表直接用新列名
ANALYTE_COL_MAP = {
    "Logp": "C-LogP", "LogP": "C-LogP",
    "Exact_mass": "C-精确分子量",
    "Hydrogen_bond_acceptor_count": "C-HBA",
    "Hydrogen_bond_donor_count": "C-HBD",
    "Strongest_acidic_pKa": "C-pka-A",
    "Strongert_basic_pKa": "C-pka-B",
    "Isoelectric_point": "C-Isoelectricpoint",
    "Ring_count": "C-芳香环数Aromaticringcount",
    "CAS_number": "C-CAS号",
    "分子量": "C-精确分子量",
    # 兼容表头变体（无 C- 前缀的裸名）→ 供 resolve_column 反向匹配
    "HBA": "C-HBA", "HBD": "C-HBD", "TPSA": "C-TPSA",
    "pka-A": "C-pka-A", "pka-B": "C-pka-B", "等电点": "C-Isoelectricpoint",
    "精确分子量": "C-精确分子量", "CAS": "C-CAS号",
    "芳香环数Aromaticringcount": "C-芳香环数Aromaticringcount",
}


def resolve_column(col: str, available: list) -> Optional[str]:
    """把（可能是旧名/新名/裸名变体的）展示列解析为表中实际存在的列；找不到返回 None。
    兼容：直接命中 → MAP 正向 → MAP 反向（请求 C-LogP 而表中只有 LogP）→ 大小写兜底。"""
    if col in available:
        return col
    mapped = ANALYTE_COL_MAP.get(col)
    if mapped and mapped in available:
        return mapped
    for a, b in ANALYTE_COL_MAP.items():
        if b == col and a in available:
            return a
    _avl_l = {str(x).strip().lower(): x for x in available}
    return _avl_l.get(str(col).strip().lower())


# 标准缩写 → 信息库全名（PFAS 系列等常用别名；大小写不敏感）。
# 若信息表本身带 缩写/Abbreviation/Alias 列，会自动建表并优先于内置别名。
ANALYTE_ALIASES = {
    "PFBA": "heptafluorobutanoic acid",
    "PFPA": "nonafluoropentanoic acid",
    "PFHxA": "undecafluorohexanoic acid",
    "PFHpA": "tridecafluoroheptanoic acid",
    "PFOA": "pentadecafluorooctanoic acid",
    "PFNA": "heptadecafluorononanoic acid",
    "PFDA": "nonadecafluorodecanoic acid",
    "PFUnDA": "henicosafluoroundecanoic acid",
    "PFDoDA": "tridecafluorododecanoic acid",
    "PFTriDA": "pentadecafluorotetradecanoic acid",
    "PFTeDA": "heptadecafluorotetradecanoic acid",
    "PFBS": "nonafluorobutane-1-sulfonic acid",
    "PFHxS": "tridecafluorohexane-1-sulfonic acid",
    "PFHpS": "pentadecafluoroheptane-1-sulfonic acid",
    "PFOS": "heptadecafluorooctane-1-sulfonic acid",
    "PFDS": "perfluorodecane sulfonic acid",
}
_ALIAS_COL_KEYS = {"缩写", "英文缩写", "简称", "abbreviation", "alias", "abbr"}


def _alias_map() -> dict:
    """别名 → 信息库主键名。自动读取信息表的缩写列（若有）+ 内置 PFAS 别名。"""
    m = {}
    df = _load_analyte_chemical()
    if df is not None and len(df.columns):
        for c in df.columns:
            if str(c).strip().lower() in _ALIAS_COL_KEYS:
                for idx, v in zip(df.index, df[c]):
                    if v not in (None, "", "None"):
                        m[_norm_analyte_name(str(v))] = str(idx)
                break
    for k, v in ANALYTE_ALIASES.items():
        m.setdefault(_norm_analyte_name(k), v)
    return m


def _is_name_col(col) -> bool:
    """名称类列：列名含 name/名（Input-englishname、物质名IUPACname、商品名Traditionalname…）；
    排除图片/数值列（结构式、photo、solubility photo 等不含 name/名）。"""
    s = str(col).strip().lower()
    return ("name" in s) or ("名" in s)


@lru_cache(maxsize=1)
def _name_forms_map() -> dict:
    """norm(任一名称形式) → 该物质行的全部名称形式（元组，去重保序）。
    来源：信息表主键 + 各名称类列的值 + 内置 ANALYTE_ALIASES（缩写→全名补充）。
    使 缩写/英文名/商品名 任一形式都能互查。"""
    forms = {}
    df = _load_analyte_chemical()
    if df is not None and len(df):
        for idx in df.index:
            cands = [str(idx)]
            for c in df.columns:
                if _is_name_col(c):
                    v = df.loc[idx, c]
                    if v not in (None, "", "None"):
                        cands.append(str(v))
            seen, uniq = set(), []
            for c in cands:
                n = _norm_analyte_name(c)
                if n and n not in seen:
                    seen.add(n)
                    uniq.append(c)
            for c in uniq:
                forms.setdefault(_norm_analyte_name(c), tuple(uniq))
    for k, v in ANALYTE_ALIASES.items():
        nk, nv = _norm_analyte_name(k), _norm_analyte_name(v)
        entry = forms.get(nk, ())
        forms[nk] = tuple(dict.fromkeys(list(entry) + [k, v]))
        entry2 = forms.get(nv, ())
        forms[nv] = tuple(dict.fromkeys(list(entry2) + [k, v]))
    return forms


def resolve_analyte_alias(analyte: str) -> str:
    """把（可能是缩写的）物质名解析为信息库主键名；无别名时原样返回（归一化后）。"""
    n = _norm_analyte_name(str(analyte))
    hit = _alias_map().get(n)
    return hit if hit else str(analyte)


def _norm_analyte_name(s: str) -> str:
    """归一化物质名：5.Fluorouracil ↔ 5-Fluorouracil ↔ 5-Fluorouracil-1 → 5-fluorouracil"""
    if s is None:
        return ""
    s = str(s).strip().lower()
    # 去尾部离子对号
    import re
    s = re.sub(r"-\d+$", "", s)
    # . 与 - 互通
    s = s.replace(".", "-")
    return s


@lru_cache(maxsize=4)
def _analyte_basic_cached() -> Optional[pd.DataFrame]:
    df = _load_analyte_chemical()
    if df is None:
        return None
    return df.reset_index()


def get_analyte_basic() -> Optional[pd.DataFrame]:
    """返回分析物基本信息表（chemical 表，重置索引便于展示；lru 缓存，信息库刷新时清除）"""
    cached = _analyte_basic_cached()
    return None if cached is None else cached.copy()


@lru_cache(maxsize=8192)
def _analyte_info_cached(analyte: str) -> Optional[pd.Series]:
    df = _load_analyte_chemical()
    if df is None:
        return None
    target = _norm_analyte_name(analyte)
    for idx in df.index:
        if _norm_analyte_name(str(idx)) == target:
            return df.loc[idx]
    # 名形映射：查询名可能是该行 缩写/IUPAC名/商品名 中的任一种（主键只是其一）
    for cand in _name_forms_map().get(target, ()):
        c_target = _norm_analyte_name(cand)
        for idx in df.index:
            if _norm_analyte_name(str(idx)) == c_target:
                return df.loc[idx]
    # 兜底：旧内置别名
    alias = _alias_map().get(target)
    if alias:
        a_target = _norm_analyte_name(alias)
        for idx in df.index:
            if _norm_analyte_name(str(idx)) == a_target:
                return df.loc[idx]
    return None


def get_analyte_info(analyte: str) -> Optional[pd.Series]:
    """按物质名查基本信息（自动归一化匹配；lru 缓存，信息库刷新时清除）"""
    cached = _analyte_info_cached(str(analyte))
    return None if cached is None else cached.copy()


@lru_cache(maxsize=1)
def _load_wide_table(sheet: str) -> Optional[pd.DataFrame]:
    """加载 LogD 或 isoelectric 宽表。返回长表格式：name, prop, ph, value"""
    path = config.ANALYTE_FILE
    if not path.exists():
        return None
    try:
        df = pd.read_excel(path, sheet_name=sheet)
    except Exception:
        return None
    df.columns = [str(c).strip() for c in df.columns]
    if len(df.columns) < 3:
        return None
    name_col = df.columns[0]      # name
    prop_col = df.columns[1]      # logD / Charge
    ph_cols = df.columns[2:]      # pH 值列名
    # 转长表
    records = []
    for _, row in df.iterrows():
        name = _clean_val(row[name_col])
        prop = _clean_val(row[prop_col])
        if name is None:
            continue
        for ph_col in ph_cols:
            try:
                ph = float(str(ph_col).strip())
            except ValueError:
                continue
            val = _try_float(row[ph_col])
            if val is not None:
                records.append({"name": name, "prop": prop, "ph": ph, "value": val})
    if not records:
        return None
    return pd.DataFrame(records)


@lru_cache(maxsize=8192)
def get_property_by_ph(analyte: str, prop: str, ph: float) -> Optional[float]:
    """
    按 pH 查宽表属性（logD/charge/solubility 等任意随 pH 变化的属性）。
    自动在 analyte.xlsx 的所有 pH 宽表 sheet 中按属性名匹配（大小写不敏感）。
    带 lru_cache：自定义看板逐行查询时相同 (物质,属性,pH) 直接命中，不再重复过滤长表。
    返回最接近给定 pH 的值；查不到返回 None
    """
    prop_l = str(prop).strip().lower()
    props = list_ph_properties()
    if prop_l not in props:
        return None
    df = _load_wide_table(props[prop_l]["sheet"])
    if df is None:
        return None
    def _match(t):
        return df[df["name"].map(lambda x: _norm_analyte_name(str(x)) == t)]

    target = _norm_analyte_name(analyte)
    sub = _match(target)
    if len(sub) == 0:
        # 查询名可能是缩写/商品名等 → 用名形映射展开为该物质全部名称形式再试
        for cand in _name_forms_map().get(target, ()):
            sub = _match(_norm_analyte_name(cand))
            if len(sub):
                break
    if len(sub) == 0:
        alias = _alias_map().get(target)
        if alias:
            sub = _match(_norm_analyte_name(alias))
    if len(sub) == 0:
        return None
    # 找最接近 ph 的行
    sub = sub.copy()
    sub["diff"] = (sub["ph"] - ph).abs()
    best = sub.sort_values("diff").iloc[0]
    return float(best["value"])


@lru_cache(maxsize=1)
def list_ph_properties() -> dict:
    """
    扫描 analyte.xlsx，自动识别所有"pH 宽表"sheet（name | 属性名 | pH 0~14 列）。
    返回 {属性名小写: {display: 原属性名, sheet: sheet名}}。
    新增 pH 类 sheet（如 solubility）无需改代码，刷新信息库后自动可用。
    """
    path = config.ANALYTE_FILE
    if not path.exists():
        return {}
    try:
        xl = pd.ExcelFile(path)
    except Exception:
        return {}
    out = {}
    for sn in xl.sheet_names:
        if sn.strip().lower() in ("chemical",):
            continue  # 基础性质表不是 pH 宽表
        df = _load_wide_table(sn)
        if df is None or len(df) == 0:
            continue
        for p in df["prop"].dropna().unique():
            key = str(p).strip().lower()
            if key and key not in out:
                out[key] = {"display": str(p).strip(), "sheet": sn}
    return out


def _photo_dirs(kind: str, default_dir: Path) -> list:
    """返回候选图片目录列表：用户自定义目录优先，默认目录兜底"""
    dirs = []
    try:
        custom = settings.get_photo_dir(kind)
        if custom and Path(custom).is_dir():
            dirs.append(Path(custom))
    except Exception:
        pass
    dirs.append(default_dir)
    # 去重
    seen, out = set(), []
    for d in dirs:
        if d not in seen:
            seen.add(d)
            out.append(d)
    return out


def _stem_norm_photo(stem: str) -> str:
    """图片文件名归一：小写、去 -structures/-结构式 等后缀词与空白"""
    st = str(stem).lower().strip()
    for suf in ("-structures", "_structures", "-结构式", "结构式"):
        st = st.replace(suf, "")
    return st.strip("-_ ")


@lru_cache(maxsize=2048)
def _analyte_photo_cached(analyte: str, _dirs_sig: tuple) -> Optional[Path]:
    """物质图片查找（按目录签名缓存：新增/改名图片自动失效）。
    候选 = 查询名 + 信息行全部名称类列的值（缩写/IUPAC名/商品名互为候选），
    每个候选先大小写不敏感精确匹配，再双向包含模糊匹配。"""
    dirs = [Path(d) for d, _m in _dirs_sig]
    candidates = [analyte]
    info = get_analyte_info(analyte)
    if info is not None:
        for c, v in info.items():
            if _is_name_col(c) and v not in (None, "", "None"):
                candidates.append(str(v))
        if "photo" in info.index and info["photo"]:
            candidates.append(str(info["photo"]))
    # 去重保序（归一化后）
    seen, cands = set(), []
    for c in candidates:
        n = _norm_analyte_name(c)
        if n and n not in seen:
            seen.add(n)
            cands.append(c)
    # 1) 精确（大小写不敏感）
    for d in dirs:
        if not d.exists():
            continue
        by_stem = {}
        for p in d.iterdir():
            if p.suffix.lower() in (".png", ".jpg", ".jpeg"):
                by_stem[p.stem.lower()] = p
        for cand in cands:
            hit = by_stem.get(cand.lower())
            if hit:
                return hit
    # 2) 模糊：双向包含（"PFBA" 不会误含，但全名 ↔ 文件名大小写/后缀差异能容）
    for d in dirs:
        if not d.exists():
            continue
        for cand in cands:
            cl = _stem_norm_photo(cand)
            if not cl:
                continue
            for p in sorted(d.iterdir()):
                if p.suffix.lower() not in (".png", ".jpg", ".jpeg"):
                    continue
                pl = _stem_norm_photo(p.stem)
                if cl in pl or pl in cl:
                    return p
    return None


def get_analyte_photo(analyte: str) -> Optional[Path]:
    """查物质结构式图片（用户自定义目录优先）。
    名称三形式通用：查询名/主键列/IUPAC名/商品名 任一形式命中的文件都可匹配。"""
    dirs = _photo_dirs("analyte", config.ANALYTE_PHOTO_DIR)
    sig = tuple((str(d), (d.stat().st_mtime if d.exists() else 0)) for d in dirs)
    return _analyte_photo_cached(str(analyte), sig)


# ============ SLM 信息库 ============

# 主SLM来源工作表
SLM_MAIN_SHEETS = ["LF-液膜类", "DES-深共晶类"]
SLM_ADDITIVE_SHEETS = ["IL-离子液体类", "O-其他"]


@lru_cache(maxsize=4)
def _load_slm_sheet(sheet: str) -> Optional[pd.DataFrame]:
    path = config.SLM_FILE
    if not path.exists():
        return None
    try:
        df = pd.read_excel(path, sheet_name=sheet)
    except Exception:
        return None
    df.columns = [str(c).strip() for c in df.columns]
    for c in df.columns:
        df[c] = df[c].map(_clean_val)
    # 主键：试剂名称
    if "试剂名称" not in df.columns:
        return None
    df = df.set_index("试剂名称")
    df["_sheet"] = sheet  # 记录来源
    return df


def get_slm_library() -> dict:
    """返回 {main_names: [...], additive_names: [...]}"""
    main_names, add_names = [], []
    for s in SLM_MAIN_SHEETS:
        df = _load_slm_sheet(s)
        if df is not None:
            main_names.extend([str(x) for x in df.index if x])
    for s in SLM_ADDITIVE_SHEETS:
        df = _load_slm_sheet(s)
        if df is not None:
            add_names.extend([str(x) for x in df.index if x])
    return {"main_names": main_names, "additive_names": add_names}


_MEASURE_RE_CACHE = {}


def _measure_re():
    """含量/浓度后缀正则（带缓存）："名 (5%)"、"名 (10 mM)"、"名 (0.2 μM)" 等均可剥离。
    仅匹配 数字+可选单位 的括号尾缀——试剂名自带的缩写括号（如 "(BMDH)"、"(NPPE)"）不受影响。"""
    import re as _re
    key = "rx"
    if key not in _MEASURE_RE_CACHE:
        _MEASURE_RE_CACHE[key] = _re.compile(
            r"\s*[(（]\s*(\d+(?:\.\d+)?\s*(?:%|mM|μM|uM|µM|M|mg/L|g/L|μg/L|ug/L|mg·L-1)?)\s*[)）]\s*$",
            _re.IGNORECASE)
    return _MEASURE_RE_CACHE[key]


def _strip_ratio(name: str) -> str:
    """剥离名称尾部的含量/浓度后缀（"NPOE (80%)"→"NPOE"，"DEHP (10 mM)"→"DEHP"），
    用于信息库查找。只剥 数字+单位 形态的尾缀，不影响 "(BMDH)" 等名称自带括号。"""
    import re as _re
    return _measure_re().sub("", str(name)).strip()


def split_name_ratio(name: str) -> tuple:
    """把 "NPOE (80%)" 拆成 ("NPOE", "80%")、"DEHP (10 mM)" 拆成 ("DEHP", "10 mM")；
    无后缀时返回 (原名, None)。第二元素为字符串（自由单位）。"""
    import re as _re
    m = _measure_re().search(str(name).rstrip())  # 正则带 $ 尾锚，必须用 search（match 锚定串首恒失败）
    if m:
        return _strip_ratio(name), m.group(1).strip()
    return str(name), None


def get_slm_info(name: str) -> Optional[pd.DataFrame]:
    """按试剂名查信息（在所有 SLM 表里找）。返回单行 DataFrame 或 None。
    名称含比例后缀（"NPOE (80%)"）时自动剥离后查找。"""
    lookup = _strip_ratio(name)
    for s in SLM_MAIN_SHEETS + SLM_ADDITIVE_SHEETS:
        df = _load_slm_sheet(s)
        if df is None:
            continue
        if lookup in df.index:
            row = df.loc[[lookup]]
            return row.reset_index()
    return None


def get_slm_all() -> Optional[pd.DataFrame]:
    """返回所有 SLM 合并表（带 _sheet 来源列）"""
    frames = []
    for s in SLM_MAIN_SHEETS + SLM_ADDITIVE_SHEETS:
        df = _load_slm_sheet(s)
        if df is not None:
            frames.append(df.reset_index())
    if not frames:
        return None
    return pd.concat(frames, ignore_index=True)


def get_slm_photo(name: str) -> Optional[Path]:
    """查 SLM 结构图（用户自定义目录优先）。
    匹配顺序：英文名（表格第二列，照片文件基本按英文名命名）→ 试剂名称；
    每个候选先精确匹配（含 .png/.jpg），再模糊匹配（忽略 -structures 等后缀、双向包含）。
    """
    dirs = _photo_dirs("slm", config.SLM_PHOTO_DIR)
    # 候选名：英文名优先，其次试剂名称（均剥离比例后缀）
    candidates = []
    name = _strip_ratio(name)
    info = get_slm_info(name)
    if info is not None and "英文名" in info.columns:
        eng = info.iloc[0]["英文名"]
        if eng:
            candidates.append(str(eng))
    candidates.append(name)
    # 去重保序
    seen, cands = set(), []
    for c in candidates:
        if c and c not in seen:
            seen.add(c)
            cands.append(c)

    def _stem_norm(stem: str) -> str:
        # 去掉常见图片后缀词（-structures / -结构式 / _structures）与空白
        st = stem.lower().strip()
        for suf in ("-structures", "_structures", "-结构式", "结构式"):
            st = st.replace(suf, "")
        return st.strip("-_ ")

    # 1) 精确文件名
    for d in dirs:
        for cand in cands:
            for ext in [".png", ".jpg", ".jpeg"]:
                p = d / f"{cand}{ext}"
                if p.exists():
                    return p
    # 2) 模糊：双向包含（英文名 "Anethole" ↔ 文件 "Anethole-structures"）
    for d in dirs:
        if not d.exists():
            continue
        for cand in cands:
            cl = _stem_norm(cand)
            if not cl:
                continue
            for p in sorted(d.iterdir()):
                if p.suffix.lower() not in (".png", ".jpg", ".jpeg"):
                    continue
                pl = _stem_norm(p.stem)
                if cl in pl or pl in cl:
                    return p
    return None


# ---------- 自检 ----------
if __name__ == "__main__":
    print("===== analyte 基本信息表 =====")
    df = get_analyte_basic()
    if df is not None:
        print(f"行数: {len(df)}, 列: {list(df.columns)[:10]}...")
        print(df.head(3).to_string())
    print("\n===== Methotrexate 信息 =====")
    info = get_analyte_info("Methotrexate")
    if info is not None:
        print("Logp:", info.get("Logp"), "HBA:", info.get("Hydrogen_bond_accept"))
    print("\n===== LogD @ pH 7.4 =====")
    for a in ["Methotrexate", "Cyclophosphamide", "5-Fluorouracil"]:
        v = get_property_by_ph(a, "logD", 7.4)
        print(f"  {a} logD@7.4 = {v}")
    print("\n===== Charge @ pH 7.4 =====")
    for a in ["Methotrexate", "Cyclophosphamide"]:
        v = get_property_by_ph(a, "charge", 7.4)
        print(f"  {a} charge@7.4 = {v}")
    print("\n===== SLM 库 =====")
    lib = get_slm_library()
    print("主SLM数:", len(lib["main_names"]), "添加剂数:", len(lib["additive_names"]))
    print("主SLM示例:", lib["main_names"][:3])
    print("\n===== 三苯基氧膦 信息 =====")
    si = get_slm_info("三苯基氧膦")
    if si is not None:
        print(si[["试剂名称", "英文名", "LogP", "分子量"]].to_string() if "LogP" in si.columns else si.iloc[0][:5])
    print("\n===== 照片查找 =====")
    for a in ["Methotrexate", "5-Fluorouracil"]:
        p = get_analyte_photo(a)
        print(f"  {a}: {p}")
