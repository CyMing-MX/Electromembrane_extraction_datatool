# -*- coding: utf-8 -*-
"""
EME 实验数据管理系统 - SQLite 数据层

表结构：
  experiments   每次实验一行（日期/组数/元数据/QC名/SLM/备注）
  metrics       每物质×每组的计算结果（回收率/转移率/RSD）
  qc_summary    每物质×每次实验的标准品/空白均值与RSD（质控用）
  rt_records    保留时间原始记录（质控出图用）

设计：本模块只负责建表、读写；计算逻辑在 calculator.py。
"""
from __future__ import annotations
import json
import re
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd

try:
    from . import config
except ImportError:
    import config


SCHEMA = """
CREATE TABLE IF NOT EXISTS experiments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    folder TEXT UNIQUE,            -- 实验文件夹名（如 20260809-12M）
    exp_date TEXT,                 -- YYYYMMDD
    temperature TEXT,
    humidity TEXT,
    column_pressure TEXT,          -- 液相平衡柱压
    Va REAL,
    Vd REAL,
    use_sample INTEGER,            -- 0/1
    qc_names TEXT,                 -- JSON: {std_acceptor, std_donor, blk_acceptor, blk_donor}
    exp_type TEXT,                 -- 实验目的分类：slm/donor/acceptor/time/voltage/all/custom
    current_sheets TEXT,           -- 电流记录映射 JSON：[{file, sheet}]（用户在新增实验时选择）
    excluded_samples TEXT,         -- 已剔除样本 JSON dict（samples/std_acceptor/std_donor）
    qc_combos TEXT,                -- 多标准品组合 JSON list [{name, std_acceptor, std_donor, blk_acceptor, blk_donor}]
    sample_format TEXT,            -- 样品名称格式（本实验覆盖；NULL=跟随全局设置）
    purpose TEXT,                  -- 实验说明
    n_groups INTEGER,
    n_replicates INTEGER,
    data_xlsx TEXT,                -- 生成的DATA.xlsx 路径
    created_at TEXT
);

-- 组级萃取条件（每组M1~M16各自的条件；基准值也存为一行 grp=0）
CREATE TABLE IF NOT EXISTS group_conditions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    experiment_id INTEGER NOT NULL,
    grp INTEGER NOT NULL,          -- 0=基准条件，1~16=各组
    group_label TEXT,              -- 组别名（自动生成或用户命名）
    slm_main TEXT,                 -- 主SLM名（LF/DES表）
    slm_additives TEXT,            -- SLM添加剂（JSON list，来自IL/O表）
    pH_donor REAL,                 -- 供体相pH
    donor_additives TEXT,          -- 供体相添加剂（JSON list，来自IL/O表）
    donor_composition TEXT,        -- 供体相成分
    donor_buffer_conc TEXT,        -- 供体相缓冲剂浓度
    donor_additive_conc TEXT,      -- 供体相添加剂浓度
    pH_acceptor REAL,              -- 受体相pH
    acceptor_additives TEXT,       -- 受体相添加剂（JSON list，来自IL/O表）
    acceptor_composition TEXT,     -- 受体相成分
    acceptor_buffer_conc TEXT,     -- 受体相缓冲剂浓度
    acceptor_additive_conc TEXT,   -- 受体相添加剂浓度
    extract_time REAL,             -- 萃取时间(分钟)
    extract_voltage REAL,          -- 萃取电压(V)
    group_note TEXT,               -- 备注（手动自由文本）
    note_tags TEXT,                -- 异常标签 JSON list（渗漏/烧焦/破裂等）
    well_position TEXT,            -- 该组使用的孔位 JSON list（如 ["A1","A2","A3"]；兼容旧单值"A1"）
    abnormal_wells TEXT,           -- 发生异常的孔位 JSON list（可选精细记录；空=整组孔位视为异常）
    qc_combo TEXT,                 -- 本组使用的标准品组合名（多标准品组合模式；空=默认组合）
    FOREIGN KEY(experiment_id) REFERENCES experiments(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS metrics (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    experiment_id INTEGER NOT NULL,
    exp_date TEXT,
    analyte TEXT,
    grp INTEGER,
    recovery REAL,                 -- 回收率%
    A_rsd REAL,                    -- 受体相RSD%
    S_rsd REAL,                    -- 供体相RSD%
    dm REAL,                       -- 供体相至SLM转移率%
    ma REAL,                       -- SLM至受体相转移率%
    corr REAL,
    A_mean REAL,
    S_mean REAL,
    FOREIGN KEY(experiment_id) REFERENCES experiments(id) ON DELETE CASCADE
);

-- 数据观察记录（整体或针对单次实验）
CREATE TABLE IF NOT EXISTS observations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    scope TEXT,                    -- 'global'（整体）/ 'experiment'（单次实验）
    experiment_id INTEGER,         -- scope=experiment 时有效
    exp_date TEXT,
    folder TEXT,
    text TEXT,
    created_at TEXT,
    image TEXT                     -- 图片相对 DATA_DIR 的路径（可空）
);

-- 看板偏好（持久化用户在自定义看板的配置）
CREATE TABLE IF NOT EXISTS dashboard_prefs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE,              -- 偏好名
    config TEXT,                   -- JSON 配置
    created_at TEXT
);

CREATE TABLE IF NOT EXISTS qc_summary (
    -- 每物质×每次实验的标准品/空白均值与RSD（供质控看板）
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    experiment_id INTEGER NOT NULL,
    exp_date TEXT,
    analyte TEXT,
    std_acceptor_mean REAL,
    std_acceptor_rsd REAL,
    std_donor_mean REAL,
    std_donor_rsd REAL,
    blk_acceptor_mean REAL,
    blk_donor_mean REAL,
    std_rt_mean REAL,              -- 标准品平均保留时间（质控RT用）
    std_rt_rsd REAL,
    std_width_mean REAL,           -- 标准品半峰宽均值（受体相标准品）
    std_width_rsd REAL,
    FOREIGN KEY(experiment_id) REFERENCES experiments(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS rt_records (
    -- 各样本各物质的保留时间原始值（供单次实验RT散点图）
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    experiment_id INTEGER NOT NULL,
    exp_date TEXT,
    sample TEXT,
    grp INTEGER,
    phase TEXT,
    replicate INTEGER,
    analyte TEXT,
    rt REAL,
    FOREIGN KEY(experiment_id) REFERENCES experiments(id) ON DELETE CASCADE
);
"""


def get_conn(db_path: Path | str | None = None) -> sqlite3.Connection:
    db_path = Path(db_path) if db_path else config.DB_PATH
    db_path.parent.mkdir(parents=True, exist_ok=True)
    # check_same_thread=False：Streamlit 在不同线程运行代码，需允许跨线程使用连接
    conn = sqlite3.connect(str(db_path), check_same_thread=False)
    # 开启外键：删除实验时关联表自动级联删除，防止孤儿数据
    conn.execute("PRAGMA foreign_keys=ON")
    conn.row_factory = sqlite3.Row
    conn.executescript(SCHEMA)
    _migrate(conn)
    return conn


def _migrate(conn: sqlite3.Connection):
    """旧库自动迁移：缺列自动补（CREATE TABLE IF NOT EXISTS 不改已有表结构）"""
    need_cols = {
        # 注意：dict 键不可重复——后写的同名键会覆盖前面的（曾因此漏迁移 sample_format）
        "experiments": [("current_sheets", "TEXT"), ("excluded_samples", "TEXT"),
                        ("sample_format", "TEXT"), ("qc_combos", "TEXT")],
        "observations": [("image", "TEXT")],
        "group_conditions": [("group_note", "TEXT"), ("note_tags", "TEXT"), ("well_position", "TEXT"),
                             ("abnormal_wells", "TEXT"), ("qc_combo", "TEXT")],
        "qc_summary": [("std_width_mean", "REAL"), ("std_width_rsd", "REAL")],
    }
    for table, cols in need_cols.items():
        existing = {r[1] for r in conn.execute(f"PRAGMA table_info({table})")}
        if not existing:
            continue
        for col, typ in cols:
            if col not in existing:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} {typ}")
        conn.commit()


# ---------- 实验 ----------

def upsert_experiment(conn: sqlite3.Connection, data: dict) -> int:
    """插入或更新一次实验（按 folder 唯一）。data 见 experiments 表字段。"""
    cur = conn.execute("SELECT id FROM experiments WHERE folder=?", (data["folder"],))
    row = cur.fetchone()
    data.setdefault("created_at", datetime.now().isoformat(timespec="seconds"))
    cols = ["folder","exp_date","temperature","humidity","column_pressure",
            "Va","Vd","use_sample","qc_names","exp_type","current_sheets","excluded_samples",
            "sample_format","qc_combos",
            "purpose","n_groups","n_replicates","data_xlsx","created_at"]
    if row:
        eid = row["id"]
        # excluded_samples 未显式提供时保留旧值（重新生成实验不得抹掉已保存的剔除记录）
        # created_at 同理：UPDATE 路径不覆盖原始创建时间
        upd_cols = [c for c in cols if c != "excluded_samples" or data.get("excluded_samples") is not None]
        upd_cols = [c for c in upd_cols if c != "created_at"]
        # qc_combos 未显式提供时保留旧值（重新生成不得抹掉已保存的多标准品组合）
        if data.get("qc_combos") is None:
            upd_cols = [c for c in upd_cols if c != "qc_combos"]
        if data.get("sample_format") is None:
            upd_cols = [c for c in upd_cols if c != "sample_format"]
        sets = ", ".join(f"{c}=?" for c in upd_cols)
        conn.execute(f"UPDATE experiments SET {sets} WHERE id=?",
                     tuple(data.get(c) for c in upd_cols) + (eid,))
        # 删除旧的明细，重写
        for t in ["metrics", "qc_summary", "rt_records", "group_conditions"]:
            conn.execute(f"DELETE FROM {t} WHERE experiment_id=?", (eid,))
        conn.commit()
        return eid
    else:
        ph = ", ".join("?" for _ in cols)
        cur = conn.execute(f"INSERT INTO experiments ({', '.join(cols)}) VALUES ({ph})",
                           tuple(data.get(c) for c in cols))
        conn.commit()
        return cur.lastrowid


def save_group_conditions(conn: sqlite3.Connection, experiment_id: int,
                          base_cond: dict, group_conds: dict):
    """
    保存组级萃取条件。
    base_cond  : 基准条件 dict（含 group_label，存为 grp=0）
    group_conds: {组号: 该组完整条件dict}（每组的所有字段都已被基准回填，
                 即调用方应传入"合并后"的完整条件，而非仅覆盖字段）
                 含 group_label（组别名）
    """
    cols = ["slm_main","slm_additives","pH_donor","donor_additives",
            "donor_composition","donor_buffer_conc","donor_additive_conc",
            "pH_acceptor","acceptor_additives","acceptor_composition",
            "acceptor_buffer_conc","acceptor_additive_conc",
            "extract_time","extract_voltage","group_note","note_tags","well_position","abnormal_wells",
            "qc_combo"]
    # 基准条件（grp=0）
    base_label = base_cond.get("group_label", "基准条件")
    row = [experiment_id, 0, base_label] + [base_cond.get(c) for c in cols]
    conn.execute(
        "INSERT INTO group_conditions (experiment_id, grp, group_label, " + ", ".join(cols) + ") "
        "VALUES (" + ", ".join("?" for _ in range(3 + len(cols))) + ")", row)
    # 各组（每组传入的是完整条件，已合并基准）
    for grp, cond in group_conds.items():
        label = cond.get("group_label", f"M{grp}")
        row = [experiment_id, int(grp), label] + [cond.get(c) for c in cols]
        conn.execute(
            "INSERT INTO group_conditions (experiment_id, grp, group_label, " + ", ".join(cols) + ") "
            "VALUES (" + ", ".join("?" for _ in range(3 + len(cols))) + ")", row)
    conn.commit()


def get_group_conditions(conn: sqlite3.Connection, experiment_id: int) -> pd.DataFrame:
    """读取某实验的组级条件（含 grp=0 基准）"""
    return pd.read_sql(
        "SELECT grp, group_label, slm_main, slm_additives, pH_donor, donor_additives, "
        "donor_composition, donor_buffer_conc, donor_additive_conc, "
        "pH_acceptor, acceptor_additives, acceptor_composition, "
        "acceptor_buffer_conc, acceptor_additive_conc, extract_time, extract_voltage, "
        "group_note, note_tags, well_position, abnormal_wells, qc_combo "
        "FROM group_conditions WHERE experiment_id=? ORDER BY grp", conn, params=(experiment_id,))


def parse_wells(raw) -> list:
    """解析孔位字段为 list：支持 JSON list（["A1","A2"]）、旧单值（"A1"）、逗号分隔（"A1,A2"）"""
    if raw is None:
        return []
    s = str(raw).strip()
    if not s:
        return []
    if s.startswith("["):
        try:
            v = json.loads(s)
            return [str(x) for x in v if x]
        except Exception:
            return []
    if "," in s or "，" in s:
        return [x.strip() for x in re.split(r"[,，]", s) if x.strip()]
    return [s]


def get_group_labels(conn: sqlite3.Connection, experiment_id: int) -> dict:
    """返回 {组号: 组别名}（不含基准 grp=0）"""
    df = get_group_conditions(conn, experiment_id)
    out = {}
    for _, r in df.iterrows():
        if r["grp"] != 0:
            label = r["group_label"] if r["group_label"] else f"M{int(r['grp'])}"
            out[int(r["grp"])] = label
    return out


# ---------- 看板偏好持久化 ----------

def save_pref(conn: sqlite3.Connection, name: str, config_dict: dict):
    """保存看板偏好（同名覆盖）"""
    cfg = json.dumps(config_dict, ensure_ascii=False)
    conn.execute(
        "INSERT INTO dashboard_prefs (name, config, created_at) VALUES (?,?,?) "
        "ON CONFLICT(name) DO UPDATE SET config=excluded.config, created_at=excluded.created_at",
        (name, cfg, datetime.now().isoformat(timespec="seconds")))
    conn.commit()


def load_pref(conn: sqlite3.Connection, name: str):
    """读取偏好；不存在返回 None"""
    row = conn.execute("SELECT config FROM dashboard_prefs WHERE name=?", (name,)).fetchone()
    if row is None:
        return None
    try:
        return json.loads(row["config"])
    except Exception:
        return None


def list_prefs(conn: sqlite3.Connection) -> list:
    rows = conn.execute("SELECT name FROM dashboard_prefs ORDER BY name").fetchall()
    return [r["name"] for r in rows]



def save_metrics(conn: sqlite3.Connection, experiment_id: int, exp_date: str,
                 metrics: list, analytes: list, groups: list):
    """保存计算结果（MetricRow 列表）"""
    rows = []
    for m in metrics:
        rows.append((
            experiment_id, exp_date, m.analyte, m.group,
            m.recovery, m.A_rsd, m.S_rsd, m.dm, m.ma, m.corr, m.A_mean, m.S_mean,
        ))
    conn.executemany(
        "INSERT INTO metrics (experiment_id, exp_date, analyte, grp, recovery, A_rsd, S_rsd, dm, ma, corr, A_mean, S_mean) "
        "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", rows)
    conn.commit()


def save_qc_summary(conn: sqlite3.Connection, experiment_id: int, exp_date: str,
                    res, analytes: list, default_qc: dict = None):
    """保存每物质的质控汇总（标准品/空白均值与RSD、标准品RT/半峰宽均值与RSD）。
    default_qc: {role: 标准品名}——多标准品实验按默认组合统计（跨实验质控看板口径）。"""
    from calculator import aggregate_qc
    _dq = default_qc or {}
    rows = []
    for a in analytes:
        std_a_mean, std_a_rsd, _ = aggregate_qc(res.df, "std_acceptor", a, sample_name=_dq.get("std_acceptor"))
        std_d_mean, std_d_rsd, _ = aggregate_qc(res.df, "std_donor", a, sample_name=_dq.get("std_donor"))
        blk_a_mean, _, _ = aggregate_qc(res.df, "blk_acceptor", a, sample_name=_dq.get("blk_acceptor"))
        blk_d_mean, _, _ = aggregate_qc(res.df, "blk_donor", a, sample_name=_dq.get("blk_donor"))
        # 标准品RT / 半峰宽（受体相标准品）
        _sa_name = _dq.get("std_acceptor")

        def _stat(col):
            sub = res.df[(res.df["category"] == "std_acceptor") & (res.df["analyte"] == a)]
            if _sa_name:
                sub = sub[sub["sample"] == _sa_name]
            vs = [x for x in sub[col].tolist() if x and not (isinstance(x, float) and x != x)]
            m = sum(vs)/len(vs) if vs else None
            rsd = (pd.Series(vs).std(ddof=1)/m*100) if (len(vs) >= 2 and m) else None
            return m, rsd
        std_rt_mean, std_rt_rsd = _stat("rt")
        std_width_mean, std_width_rsd = _stat("width50")
        rows.append((experiment_id, exp_date, a, std_a_mean, std_a_rsd,
                     std_d_mean, std_d_rsd, blk_a_mean, blk_d_mean,
                     std_rt_mean, std_rt_rsd, std_width_mean, std_width_rsd))
    conn.executemany(
        "INSERT INTO qc_summary (experiment_id, exp_date, analyte, std_acceptor_mean, std_acceptor_rsd, "
        "std_donor_mean, std_donor_rsd, blk_acceptor_mean, blk_donor_mean, std_rt_mean, std_rt_rsd, "
        "std_width_mean, std_width_rsd) "
        "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", rows)
    conn.commit()


def save_rt_records(conn: sqlite3.Connection, experiment_id: int, exp_date: str, res):
    """保存RT原始记录。样品名按解析时实际生效的格式解析（res.sample_format），
    不得退回全局设置——实验级格式覆盖会让 grp/phase/replicate 全部落 NULL。"""
    sub = res.df[res.df["category"] == "sample"]
    rows = []
    from parser import parse_sample_name
    _fmt = getattr(res, "sample_format", "") or ""
    for _, r in sub.iterrows():
        g, ph, rp = parse_sample_name(r["sample"], pattern=_fmt or None) or (None, None, None)
        rt = r["rt"]
        if isinstance(rt, float) and rt != rt:
            rt = None
        rows.append((experiment_id, exp_date, r["sample"], g, ph, rp, r["analyte"], rt))
    conn.executemany(
        "INSERT INTO rt_records (experiment_id, exp_date, sample, grp, phase, replicate, analyte, rt) "
        "VALUES (?,?,?,?,?,?,?,?)", rows)
    conn.commit()


# ---------- 查询 ----------

def list_experiments(conn: sqlite3.Connection) -> pd.DataFrame:
    return pd.read_sql("SELECT * FROM experiments ORDER BY exp_date DESC", conn)


def get_metrics_df(conn: sqlite3.Connection) -> pd.DataFrame:
    return pd.read_sql("SELECT * FROM metrics", conn)


def get_qc_df(conn: sqlite3.Connection) -> pd.DataFrame:
    return pd.read_sql("SELECT * FROM qc_summary", conn)


def get_rt_df(conn: sqlite3.Connection) -> pd.DataFrame:
    return pd.read_sql("SELECT * FROM rt_records", conn)


def delete_experiment(conn: sqlite3.Connection, experiment_id: int):
    # 观察记录（含已写盘的附图）一并清理：observations 无外键声明，漏删会永久留孤儿
    try:
        import os as _os
        for r in conn.execute("SELECT image FROM observations WHERE experiment_id=?", (experiment_id,)):
            if r["image"]:
                try:
                    # image 存的是相对 DATA_DIR 的路径（experiments/<folder>/xxx.png）
                    _p = DATA_DIR / r["image"] if not _os.path.isabs(r["image"]) else Path(r["image"])
                    if _p.exists():
                        _p.unlink()
                except OSError:
                    pass
        conn.execute("DELETE FROM observations WHERE experiment_id=?", (experiment_id,))
    except Exception:
        pass  # observations 表不存在等旧库场景
    for t in ["metrics", "qc_summary", "rt_records", "group_conditions"]:
        conn.execute(f"DELETE FROM {t} WHERE experiment_id=?", (experiment_id,))
    conn.execute("DELETE FROM experiments WHERE id=?", (experiment_id,))
    conn.commit()


def _load_excluded_raw(conn: sqlite3.Connection, experiment_id: int) -> dict:
    """读取 excluded_samples 原始 JSON。
    兼容两种格式：旧=list（纯样本名）；新=dict {"samples": [...], "std_acceptor": [[名, 针号]...], "std_donor": [...]}
    """
    row = conn.execute("SELECT excluded_samples FROM experiments WHERE id=?", (experiment_id,)).fetchone()
    if not row or not row["excluded_samples"]:
        return {}
    try:
        v = json.loads(row["excluded_samples"])
    except Exception:
        return {}
    if isinstance(v, dict):
        return v
    if isinstance(v, list):
        return {"samples": v}
    return {}


def _save_excluded_raw(conn: sqlite3.Connection, experiment_id: int, raw: dict):
    conn.execute("UPDATE experiments SET excluded_samples=? WHERE id=?",
                 (json.dumps(raw, ensure_ascii=False), experiment_id))
    conn.commit()


def get_excluded_samples(conn: sqlite3.Connection, experiment_id: int) -> set:
    """读取实验的已剔除样本集合；无记录返回空集"""
    return set(_load_excluded_raw(conn, experiment_id).get("samples") or [])


def get_all_excluded_samples(conn: sqlite3.Connection) -> dict:
    """一次 SQL 读出全部实验的剔除记录 {experiment_id: set(样本名)}（供汇总页用，避免逐实验查询）"""
    out = {}
    for row in conn.execute("SELECT id, excluded_samples FROM experiments"):
        if not row["excluded_samples"]:
            continue
        try:
            v = json.loads(row["excluded_samples"])
        except Exception:
            continue
        if isinstance(v, dict):
            out[int(row["id"])] = set(v.get("samples") or [])
        elif isinstance(v, list):
            out[int(row["id"])] = set(v)
    return out


def set_excluded_samples(conn: sqlite3.Connection, experiment_id: int, excluded: set):
    """保存实验的已剔除样本集合（保留已保存的标准品针剔除记录）"""
    raw = _load_excluded_raw(conn, experiment_id)
    raw["samples"] = sorted(excluded)
    _save_excluded_raw(conn, experiment_id, raw)


def get_excluded_std(conn: sqlite3.Connection, experiment_id: int) -> dict:
    """读取已剔除的标准品进样针。
    返回 {"std_acceptor": {(样名, 针号), ...}, "std_donor": {(样名, 针号), ...}}
    """
    raw = _load_excluded_raw(conn, experiment_id)
    out = {}
    for cat in ("std_acceptor", "std_donor"):
        pairs = raw.get(cat) or []
        try:
            out[cat] = {(str(n), int(i)) for n, i in pairs}
        except Exception:
            out[cat] = set()
    return out


def set_excluded_std(conn: sqlite3.Connection, experiment_id: int,
                     std_acceptor=None, std_donor=None):
    """保存标准品进样针剔除记录（覆盖对应相；保留样本剔除记录）。
    参数为 (样名, 针号) 对的集合/列表；传 None 表示该相不动。
    """
    raw = _load_excluded_raw(conn, experiment_id)
    if std_acceptor is not None:
        raw["std_acceptor"] = [[s, int(i)] for s, i in std_acceptor]
    if std_donor is not None:
        raw["std_donor"] = [[s, int(i)] for s, i in std_donor]
    _save_excluded_raw(conn, experiment_id, raw)


# ---------- 数据观察记录 ----------

def add_observation(scope: str, text: str, experiment_id=None, exp_date="", folder="",
                    image: str = ""):
    """新增观察记录。scope: 'global'（整体）/ 'experiment'（单次实验）。
    image: 图片相对 DATA_DIR 的路径（已压缩保存后的路径）。"""
    conn = get_conn()
    conn.execute(
        "INSERT INTO observations (scope, experiment_id, exp_date, folder, text, created_at, image) "
        "VALUES (?,?,?,?,?,?,?)",
        (scope, experiment_id, exp_date, folder, text,
         datetime.now().strftime("%Y-%m-%d %H:%M"), image or ""))
    conn.commit()


def save_observation_image(folder: str, exp_date: str, uploaded_file) -> str:
    """压缩并保存观察图片到 experiments/<folder>/，自动命名。
    返回相对 DATA_DIR 的路径。图片随实验文件夹走（备份/迁移/软件更新都不影响读取）。"""
    import io
    import config as _cfg
    from PIL import Image as _PILImage
    base = _cfg.DATA_DIR / "experiments" / folder
    base.mkdir(parents=True, exist_ok=True)
    img = _PILImage.open(io.BytesIO(uploaded_file.read()))
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    max_side = 1600  # 自带压缩：最长边 ≤1600px
    w, h = img.size
    if max(w, h) > max_side:
        scale = max_side / max(w, h)
        img = img.resize((max(1, int(w * scale)), max(1, int(h * scale))), _PILImage.LANCZOS)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    name = f"obsimg_{exp_date or stamp[:8]}_{stamp[9:]}"
    path = base / f"{name}.png"
    k = 1
    while path.exists():
        path = base / f"{name}_{k}.png"
        k += 1
    img.save(path, "PNG", optimize=True)
    return f"experiments/{folder}/{path.name}"


def get_observation_image_path(rel_path: str):
    """按记录里的相对路径定位图片文件；不存在返回 None。"""
    if not rel_path:
        return None
    import config as _cfg
    p = _cfg.DATA_DIR / rel_path
    return p if p.exists() else None


def update_observation(obs_id: int, text: str):
    """修改观察记录的文本内容（图片不变）。"""
    conn = get_conn()
    conn.execute("UPDATE observations SET text=? WHERE id=?", (text, obs_id))
    conn.commit()


def list_observations(scope: str = None) -> pd.DataFrame:
    """列出观察记录（scope=None 返回全部），含实验文件夹信息"""
    conn = get_conn()
    if scope:
        return pd.read_sql(
            "SELECT o.id, o.scope, o.exp_date, o.folder, o.text, o.created_at, o.image "
            "FROM observations o WHERE o.scope=? ORDER BY o.created_at DESC, o.id DESC",
            conn, params=(scope,))
    return pd.read_sql(
        "SELECT o.id, o.scope, o.exp_date, o.folder, o.text, o.created_at, o.image "
        "FROM observations o ORDER BY o.created_at DESC, o.id DESC", conn)


def delete_observation(obs_id: int):
    conn = get_conn()
    row = conn.execute("SELECT image FROM observations WHERE id=?", (obs_id,)).fetchone()
    if row and row["image"]:
        try:
            import config as _cfg
            p = _cfg.DATA_DIR / row["image"]
            if p.exists():
                p.unlink()
        except Exception:
            pass
    conn.execute("DELETE FROM observations WHERE id=?", (obs_id,))
    conn.commit()
