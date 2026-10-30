# -*- coding: utf-8 -*-
"""
EME 实验数据管理系统 - 全局配置
所有默认值、阈值、命名规则集中在此，方便用户自定义。

路径发现规则（为打包 exe 做准备）：
1. 环境变量 EME_DATA_DIR 优先
2. 依次向上查找 base 目录下的 data/（含 eme.db 或 experiments 子目录即认定）
3. 打包 exe：base = exe 所在目录；脚本运行：base = 项目根目录；再兜底 cwd
4. 都找不到 → 在 base 下创建 data/
"""
import os
import sys
from pathlib import Path


def _base_candidates() -> list:
    """按优先级返回可能的 base 目录列表"""
    cands = []
    # 打包 exe：exe 所在目录
    if getattr(sys, "frozen", False):
        cands.append(Path(sys.executable).resolve().parent)
    # 脚本运行：eme_db 的上一级（项目根）
    cands.append(Path(__file__).resolve().parent.parent)
    # 当前工作目录及其父级（exe 双击运行时 cwd 可能是任意处）
    cwd = Path.cwd().resolve()
    cands.append(cwd)
    cands.append(cwd.parent)
    # exe 所在目录（非 frozen 但从 exe 目录启动的场景）
    cands.append(Path(sys.argv[0]).resolve().parent if sys.argv and sys.argv[0] else None)
    # 去重去 None，保持顺序
    seen, out = set(), []
    for c in cands:
        if c and c not in seen:
            seen.add(c)
            out.append(c)
    return out


def _looks_like_data_dir(p: Path) -> bool:
    """判断某目录是否像本系统的 data 目录"""
    if not p.is_dir():
        return False
    return (p / "eme.db").exists() or (p / "experiments").exists() or (p / "analyte.xlsx").exists()


def _find_data_dir() -> Path:
    """自动发现 data 目录；找不到则用第一个 base 下的 data（延迟创建）"""
    # 1. 环境变量优先
    env = os.environ.get("EME_DATA_DIR")
    if env:
        return Path(env)
    # 2. 各 base 下的 data/
    for base in _base_candidates():
        d = base / "data"
        if _looks_like_data_dir(d):
            return d
    # 3. 各 base 本身就是 data（exe 与 data 平级打包的场景）
    for base in _base_candidates():
        if _looks_like_data_dir(base):
            return base
    # 4. 兜底：项目根（脚本）或 exe/cwd 目录下的 data
    fallback_base = _base_candidates()[0]
    return fallback_base / "data"


ROOT_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = _find_data_dir()
DB_PATH = DATA_DIR / "eme.db"
EXPERIMENTS_DIR = DATA_DIR / "experiments"
ANALYTE_FILE = DATA_DIR / "analyte.xlsx"
SLM_FILE = DATA_DIR / "SLM.xlsx"
PHOTO_DIR = DATA_DIR / "photo"
ANALYTE_PHOTO_DIR = PHOTO_DIR / "Analytes"
SLM_PHOTO_DIR = PHOTO_DIR / "SLM"

# ============ 解析配置 ============
# 定量离子对：默认取每个物质的第 1 个离子对（如 Methotrexate-1）
DEFAULT_ION_PAIR = 1

# 正向提取样本的正则：组号-A/S-重复号，如 1-A-1, 12-S-3
SAMPLE_NAME_PATTERN = r"^(?P<g>\d+)-(?P<ph>[AS])-(?P<r>\d+)$"

# 默认质控样/空白样名称（本次实验与默认不一致时由用户在界面输入）
DEFAULT_STD_ACCEPTOR = "Accept-STD-100ppb-1"   # 受体相标准品（两相溶液相同时也用这个）
DEFAULT_STD_DONOR = "Accept-STD-100ppb-1"      # 供体相标准品（两相溶液相同时与受体相同）
DEFAULT_BLK_ACCEPTOR = "Accept-blank-1"        # 受体相空白
DEFAULT_BLK_DONOR = "Accept-blank-1"           # 供体相空白（两相溶液相同时与受体相同）

# ============ 体积默认值 ============
DEFAULT_VA = 50.0    # 受体相体积 μL
DEFAULT_VD = 250.0   # 供体相体积 μL

# ============ 计算配置 ============
# RSD 阈值（百分比），超过则高亮
RSD_THRESHOLD = 15.0

# 转移率有效范围，超出标红
TRANSFER_RATE_MIN = 0.0
TRANSFER_RATE_MAX = 100.0

# 自定义公式可用的变量名（在 simpleeval 命名空间中暴露）
# 含义见各变量注释，calculator.py 会注入实际值
FORMULA_VARIABLES = {
    "A_mean":              "受体相峰面积均值（扣除空白前）",
    "S_mean":              "供体相峰面积均值（扣除空白前）",
    "STD_acceptor_mean":   "受体相标准品峰面积均值",
    "STD_donor_mean":      "供体相标准品峰面积均值",
    "blk_acceptor_mean":   "受体相空白峰面积均值",
    "blk_donor_mean":      "供体相空白峰面积均值",
    "Va":                  "受体相体积 μL",
    "Vd":                  "供体相体积 μL",
    "pH_donor":            "供体相 pH",
    "pH_acceptor":         "受体相 pH",
    "corr":                "校正系数 = STD_donor_mean / STD_acceptor_mean",
}

# ============ 看板默认显示的分析物性质列 ============
# 物质下钻时默认显示这几列（可由用户在界面增减）
DEFAULT_ANALYTE_DISPLAY_COLS = ["LogP", "MW"]

# ============ 文件命名约定 ============
# 实验文件夹名格式：日期-组数M，如 20260809-12M
EXPERIMENT_FOLDER_PATTERN = r"^(\d{8})-(\d+)M$"

# ============ 元数据字段定义（界面输入用）============
# (字段key, 显示名, 默认值, 是否必填)
# 文本类默认值用 str，数值类用 float
META_FIELDS = [
    # —— 时间环境 ——
    ("exp_date",       "实验日期(YYYYMMDD)", "", True),
    ("exp_time",       "实验时间",           "", False),
    ("temperature",    "温度(℃)",           "", False),
    ("humidity",       "湿度(%)",            "", False),
    # —— pH ——
    ("pH_donor",       "供体相pH",           "", False),
    ("pH_acceptor",    "受体相pH",           "", False),
    # —— 体积 ——
    ("Va",             "受体相体积μL",       DEFAULT_VA, True),
    ("Vd",             "供体相体积μL",       DEFAULT_VD, True),
    # —— 液相/萃取条件（自由文本）——
    ("lc_condition",   "液相条件",           "", False),
    ("extract_condition", "萃取条件",         "", False),
    # —— 备注 ——
    ("purpose",        "实验目的/备注",      "", False),
]
