# -*- coding: utf-8 -*-
"""
EME 实验数据管理系统 - 用户设置持久化

用户的自定义默认值（标准品名、pH、液相条件、Va/Vd 等）存到 JSON 文件，
下次启动软件时自动加载，覆盖 config.py 的内置默认值。
"""
from __future__ import annotations
import json
import re
from pathlib import Path

try:
    from . import config
except ImportError:
    import config

SETTINGS_FILE = config.DATA_DIR / "user_settings.json"

# 用户可配置的默认值字段定义：(key, 显示名, 内置默认值, 类型)
# 类型用于界面渲染：text/number/select
SETTING_FIELDS = [
    ("DEFAULT_STD_ACCEPTOR", "受体相标准品默认名", config.DEFAULT_STD_ACCEPTOR, "text"),
    ("DEFAULT_STD_DONOR",    "供体相标准品默认名", config.DEFAULT_STD_DONOR, "text"),
    ("DEFAULT_BLK_ACCEPTOR", "受体相空白默认名", config.DEFAULT_BLK_ACCEPTOR, "text"),
    ("DEFAULT_BLK_DONOR",    "供体相空白默认名", config.DEFAULT_BLK_DONOR, "text"),
    ("DEFAULT_VA",           "受体相体积 Va (μL)", config.DEFAULT_VA, "number"),
    ("DEFAULT_VD",           "供体相体积 Vd (μL)", config.DEFAULT_VD, "number"),
    ("DEFAULT_PH_DONOR",     "基准供体相 pH", 7.0, "number"),
    ("DEFAULT_PH_ACCEPTOR",  "基准受体相 pH", 3.0, "number"),
    ("DEFAULT_EXTRACT_TIME", "基准萃取时间 (分钟)", 10.0, "number"),
    ("DEFAULT_EXTRACT_VOLTAGE", "基准萃取电压 (V)", 80.0, "number"),
    ("DEFAULT_TEMPERATURE",  "默认温度 (℃)", "", "text"),
    ("DEFAULT_HUMIDITY",     "默认湿度 (%)", "", "text"),
    ("DEFAULT_COLUMN_PRESSURE", "默认液相平衡柱压", "", "text"),
    ("NOTE_TAGS", "异常备注预设标签(逗号分隔)", "渗漏,烧焦,破裂", "text"),
    # 组级条件默认值
    ("DEFAULT_DONOR_COMPOSITION", "基准供体相成分", "", "text"),
    ("DEFAULT_DONOR_BUFFER_CONC", "基准供体相缓冲剂浓度", "", "text"),
    ("DEFAULT_DONOR_ADDITIVE_CONC", "基准供体相添加剂浓度", "", "text"),
    ("DEFAULT_ACCEPTOR_COMPOSITION", "基准受体相成分", "", "text"),
    ("DEFAULT_ACCEPTOR_BUFFER_CONC", "基准受体相缓冲剂浓度", "", "text"),
    ("DEFAULT_ACCEPTOR_ADDITIVE_CONC", "基准受体相添加剂浓度", "", "text"),
]

# SETTING_FIELDS 的 key → 默认值映射（供 get_setting 兜底）
_FIELD_DEFAULTS = {f[0]: f[2] for f in SETTING_FIELDS}


_SETTINGS_CACHE = {"mtime": None, "data": {}}


def load_settings() -> dict:
    """加载用户设置；不存在则返回空 dict（调用方用 config 内置默认兜底）。
    按 mtime 缓存：get_setting 每次渲染被调用 ~10-16 次，避免每次都读盘。
    save_settings 写入后 mtime 变化，自动失效。"""
    try:
        mtime = SETTINGS_FILE.stat().st_mtime if SETTINGS_FILE.exists() else None
    except OSError:
        mtime = None
    if mtime != _SETTINGS_CACHE["mtime"]:
        data = {}
        if mtime is not None:
            try:
                with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception:
                data = {}
        _SETTINGS_CACHE["mtime"] = mtime
        _SETTINGS_CACHE["data"] = data
    return _SETTINGS_CACHE["data"]


def save_settings(settings: dict):
    """保存用户设置"""
    config.DATA_DIR.mkdir(parents=True, exist_ok=True)
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)
    # 写后立即同步缓存（防止 mtime 粒度问题读到旧值）
    try:
        _SETTINGS_CACHE["mtime"] = SETTINGS_FILE.stat().st_mtime
    except OSError:
        _SETTINGS_CACHE["mtime"] = None
    _SETTINGS_CACHE["data"] = settings


def get_setting(key: str, default=None):
    """读取某项设置；优先用户配置，其次 config.py，再其次 SETTING_FIELDS 内置默认"""
    user = load_settings()
    if key in user:
        return user[key]
    # 从 config 取
    if hasattr(config, key):
        return getattr(config, key)
    # 从 SETTING_FIELDS 取内置默认
    if key in _FIELD_DEFAULTS:
        return _FIELD_DEFAULTS[key]
    return default


def reset_settings():
    """重置为默认（删除用户配置文件）"""
    if SETTINGS_FILE.exists():
        SETTINGS_FILE.unlink()


# ============ 样品名称格式（组号/相/重复号 的识别规则）============
# 命名组约定：g=组号、ph=相位（A=受体相，S=供体相）、r=重复号。
# 全局默认存于 user_settings.json 的 SAMPLE_NAME_FORMAT；单个实验可在新增实验页临时覆盖
# （覆盖值随实验存库，之后重算/重新生成始终按该实验自己的格式）。


def get_sample_format() -> str:
    """全局样品名称正则；未设置或非法时回退 config 内置默认（组-相-重复）。"""
    v = get_setting("SAMPLE_NAME_FORMAT", "")
    v = str(v).strip() if v else ""
    if not v:
        return config.SAMPLE_NAME_PATTERN
    try:
        rx = re.compile(v)
        if not {"g", "ph", "r"}.issubset(rx.groupindex.keys()):
            return config.SAMPLE_NAME_PATTERN
        return v
    except re.error:
        return config.SAMPLE_NAME_PATTERN


SAMPLE_FORMAT_PRESETS = {
    "组-相-重复（如 1-A-1）": r"^(?P<g>\d+)-(?P<ph>[AS])-(?P<r>\d+)$",
    "相组-重复（如 A1-1）": r"^(?P<ph>[AS])(?P<g>\d+)-(?P<r>\d+)$",
}


# ============ 自定义试剂标签（用户在添加剂/试剂下拉中自由输入的名称）============
# 存于 user_settings.json 的 CUSTOM_REAGENTS；若之后 SLM 表更新了同名试剂，
# 该名称自动"转正"（从自定义清单中清除，选项由正式试剂表提供，信息按名称自然匹配）。


def get_custom_reagents() -> list:
    v = get_setting("CUSTOM_REAGENTS", [])
    return [str(x).strip() for x in v if str(x).strip()] if isinstance(v, list) else []


def add_custom_reagent(name: str):
    name = str(name).strip()
    if not name:
        return
    lst = get_custom_reagents()
    if name not in lst:
        lst.append(name)
        s = load_settings()
        s["CUSTOM_REAGENTS"] = lst
        save_settings(s)


def remove_custom_reagent(name: str):
    lst = get_custom_reagents()
    if name in lst:
        lst.remove(name)
        s = load_settings()
        s["CUSTOM_REAGENTS"] = lst
        save_settings(s)


def purge_custom_reagents(lib_names) -> list:
    """删除已与试剂表重名的自定义标签（自动转正）。返回剩余清单。"""
    lib_names = set(lib_names)
    remain = [n for n in get_custom_reagents() if n not in lib_names]
    s = load_settings()
    s["CUSTOM_REAGENTS"] = remain
    save_settings(s)
    return remain


# ============ 配色偏好（记住用户上次选择的颜色渐变）============
CMAP_OPTIONS = ["RdYlGn", "viridis", "plasma", "Blues", "RdBu", "YlGnBu", "coolwarm"]


def get_cmap_pref(default: str = "RdYlGn") -> str:
    """上次使用的配色名；非法值回退 default"""
    v = get_setting("CMAP_PREFERENCE", default)
    return v if v in CMAP_OPTIONS else default


def save_cmap_pref(name: str):
    """保存配色偏好（在下拉框 on_change 回调中调用）"""
    if name in CMAP_OPTIONS:
        s = load_settings()
        s["CMAP_PREFERENCE"] = name
        save_settings(s)


# ============ 指标中文全称映射（全局统一）============
# 数据库列名 → 中文全称（用于所有 web 展示）
METRIC_LABELS = {
    "recovery": "回收率%",
    "dm":       "供体相至SLM转移率%",
    "ma":       "SLM至受体相转移率%",
    "A_mean":   "受体峰面积平均值",
    "S_mean":   "供体峰面积平均值",
    "A_rsd":    "受体峰面积RSD%",
    "S_rsd":    "供体峰面积RSD%",
    "corr":     "校正系数",
}

# ============ 详情展示列自定义（物质/SLM 关键信息）============

# 分析物详情侧栏默认展示的列（分子描述符等）
# 兼容 analyte.xlsx 新版列名（C- 前缀）；旧列名会经 reference.ANALYTE_COL_MAP 自动映射
DEFAULT_ANALYTE_DISPLAY = ["C-LogP", "C-精确分子量", "C-HBA", "C-HBD",
                           "C-pka-A", "C-pka-B", "C-Isoelectricpoint",
                           "C-TPSA", "Logp", "分子量"]

# SLM 详情卡片默认展示的列
DEFAULT_SLM_DISPLAY = ["英文名", "缩写", "CAS号", "分类", "分子式", "分子量",
                       "LogP", "HLB", "HBD", "HBA", "熔点"]


def get_display_columns(kind: str) -> list:
    """
    读取用户自定义的详情展示列。
    kind: 'analyte' 或 'slm'
    用户设置存 SETTINGS_FILE 的 'display_cols_analyte' / 'display_cols_slm'（列表）
    未设置时返回内置默认。
    """
    user = load_settings()
    key = f"display_cols_{kind}"
    if key in user and isinstance(user[key], list):
        return user[key]
    return DEFAULT_ANALYTE_DISPLAY if kind == "analyte" else DEFAULT_SLM_DISPLAY


def save_display_columns(kind: str, cols: list):
    """保存用户自定义的展示列"""
    user = load_settings()
    user[f"display_cols_{kind}"] = cols
    save_settings(user)

# 反向映射（中文 → 列名），用于界面选择后转回列名查询
LABEL_TO_METRIC = {v: k for k, v in METRIC_LABELS.items()}

# ============ 自定义公式存取 ============

# 自定义公式最多条数
MAX_CUSTOM_FORMULAS = 3


def get_custom_formulas() -> list:
    """读取用户自定义公式列表：[{name, expr}]；不存在返回 []"""
    user = load_settings()
    v = user.get("custom_formulas", [])
    return v if isinstance(v, list) else []


def save_custom_formulas(formulas: list):
    """保存自定义公式列表（最多 MAX_CUSTOM_FORMULAS 条）"""
    user = load_settings()
    user["custom_formulas"] = formulas[:MAX_CUSTOM_FORMULAS]
    save_settings(user)


# ============ 摩尔浓度试剂表（分子量/密度，可编辑）============

# 内置默认表（用户可在计算器页编辑/新增，自定义部分存 user_settings.json）
DEFAULT_REAGENT_TABLE = [
    {"名称": "甲酸", "分子量(g/mol)": 46.03, "密度(g/mL)": 1.22},
    {"名称": "三氟乙酸", "分子量(g/mol)": 114.02, "密度(g/mL)": 1.49},
    {"名称": "磷酸", "分子量(g/mol)": 97.99, "密度(g/mL)": 1.89},
    {"名称": "乙酸", "分子量(g/mol)": 60.05, "密度(g/mL)": 1.05},
    {"名称": "盐酸", "分子量(g/mol)": 36.46, "密度(g/mL)": 1.18},
    {"名称": "乳酸", "分子量(g/mol)": 90.08, "密度(g/mL)": 1.21},
]


def get_reagent_table() -> list:
    """读取摩尔试剂表：自定义表覆盖内置（同名条目以自定义为准）"""
    user = load_settings()
    custom = user.get("reagent_table", [])
    if not isinstance(custom, list):
        custom = []
    merged = {r["名称"]: dict(r) for r in DEFAULT_REAGENT_TABLE}
    for r in custom:
        if isinstance(r, dict) and r.get("名称"):
            merged[r["名称"]] = r
    return list(merged.values())


def save_reagent_table(rows: list):
    """保存试剂表（仅保存与内置不同的自定义行？——简化：整表保存）"""
    user = load_settings()
    clean = []
    for r in rows:
        if isinstance(r, dict) and r.get("名称"):
            clean.append({"名称": str(r["名称"]),
                          "分子量(g/mol)": float(r.get("分子量(g/mol)") or 0),
                          "密度(g/mL)": float(r.get("密度(g/mL)") or 0)})
    user["reagent_table"] = clean
    save_settings(user)


# ============ 图片文件夹自定义 ============

def get_photo_dir(kind: str):
    """
    读取用户自定义图片文件夹路径（Path 或 None）。
    kind: 'analyte' 或 'slm'
    """
    user = load_settings()
    p = user.get(f"photo_dir_{kind}")
    from pathlib import Path as _P
    return _P(p) if p else None


def save_photo_dir(kind: str, path: str):
    """保存自定义图片文件夹路径（空字符串=恢复默认）"""
    user = load_settings()
    user[f"photo_dir_{kind}"] = path.strip()
    save_settings(user)


def get_note_tags() -> list:
    """读取异常备注预设标签列表"""
    raw = get_setting("NOTE_TAGS", "渗漏,烧焦,破裂")
    return [t.strip() for t in str(raw).split(",") if t.strip()]


# 参考线默认值（回收率与转移率分开）
_REF_DEFAULTS = {
    "recovery": (20.0, 40.0),
    "transfer": (None, None),  # 转移率默认不启用具体数值参考线
}


def get_ref_lines(kind: str = "recovery") -> tuple:
    """读取某类指标的统计参考线默认值；返回 (ref1, ref2)，未设置时 ref 为 None"""
    user = load_settings()
    d1, d2 = _REF_DEFAULTS.get(kind, (None, None))
    r1 = user.get(f"ref_line_{kind}_1", d1)
    r2 = user.get(f"ref_line_{kind}_2", d2)
    def _f(v):
        try:
            return float(v)
        except (TypeError, ValueError):
            return None
    return _f(r1), _f(r2)


def save_ref_lines(ref1, ref2, kind: str = "recovery"):
    """保存某类指标的统计参考线默认值（None 表示不启用该线）"""
    user = load_settings()
    user[f"ref_line_{kind}_1"] = ref1
    user[f"ref_line_{kind}_2"] = ref2
    save_settings(user)
