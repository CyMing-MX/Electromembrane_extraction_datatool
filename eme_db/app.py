# -*- coding: utf-8 -*-
"""
EME 实验数据管理系统 - Streamlit 主应用
启动：双击 run.bat 或 streamlit run app.py
"""
from __future__ import annotations
import sys
import re
import json
from pathlib import Path
from datetime import date, datetime

sys.path.insert(0, str(Path(__file__).resolve().parent))

import streamlit as st
from functools import lru_cache
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

import config
import parser as P
import calculator as C
import models as M
import excel_builder as EB
import dashboard as D
import naming as N
import settings as S
import current as CUR

try:
    import reference as R
    REF_OK = True
except Exception:
    REF_OK = False


st.set_page_config(page_title="EME 实验数据管理", page_icon="🧪", layout="wide",
                   initial_sidebar_state="expanded")

# ---------- 全局美化样式（不触碰顶栏/侧栏可见性，避免侧边栏打不开的问题）----------
st.markdown("""
<style>
    /* 全局字体与背景 */
    .stApp {font-family: "Microsoft YaHei", "PingFang SC", sans-serif;}
    .block-container {padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1500px;}
    /* 主标题 */
    h1 {font-weight: 700; letter-spacing: .5px;}
    /* 按钮：圆角+主色 */
    .stButton > button {border-radius: 8px; border: 1px solid #d0d7e2; font-weight: 500;}
    .stButton > button[kind="primary"] {background: #2f6fed; border-color: #2f6fed; color: #fff;}
    .stButton > button[kind="primary"]:hover {background: #1e5bd8;}
    /* 指标卡 */
    [data-testid="stMetric"] {background: #ffffff; border: 1px solid #e6ebf2; border-radius: 10px;
                              padding: 12px 16px; box-shadow: 0 1px 3px rgba(0,0,0,.04);}
    /* expander 圆角 */
    [data-testid="stExpander"] {border: 1px solid #e6ebf2; border-radius: 10px; overflow: hidden;}
    /* 标签页 */
    .stTabs [data-baseweb="tab-list"] {gap: 4px;}
    .stTabs [data-baseweb="tab"] {border-radius: 8px 8px 0 0; padding: 6px 14px;}
    /* 数据表标题行 */
    .stDataFrame {border-radius: 8px; overflow: hidden;}
    /* 提示条 */
    [data-testid="stAlert"] {border-radius: 10px;}
</style>
""", unsafe_allow_html=True)

EXP_TYPES = {
    "测试SLM（支撑液膜）": {"key": "slm", "hint": "各组使用不同 SLM（主+添加剂），其他条件保持一致"},
    "测试供体相": {"key": "donor", "hint": "各组供体相 pH / 添加剂不同，其他条件一致"},
    "测试受体相": {"key": "acceptor", "hint": "各组受体相 pH / 添加剂不同，其他条件一致"},
    "测试萃取时间": {"key": "time", "hint": "各组萃取时间不同，其他条件一致"},
    "测试萃取电压": {"key": "voltage", "hint": "各组萃取电压不同，其他条件一致"},
    "综合优化（多因素）": {"key": "all", "hint": "多个因素同时变化；组别名按 Mul- 前缀自动生成（含电压和时间）"},
}

# 指标中文全称（统一使用，避免缩写）
METRIC_LABELS = S.METRIC_LABELS
LABEL_TO_METRIC = S.LABEL_TO_METRIC


@st.cache_resource
def get_db():
    return M.get_conn()


# ---------- 共享解析缓存 ----------
# st.tabs 的所有标签页内容在每次交互时都会重新执行，任何未缓存的 parse_experiment
# 调用都会让整页卡顿。这里按 (文件夹, 质控名, txt文件mtime) 统一缓存；
# 返回的是缓存副本，调用方可安全修改 res.df。

def _folder_sig(folder: str) -> tuple:
    d = config.EXPERIMENTS_DIR / folder
    try:
        return tuple(sorted((p.name, p.stat().st_mtime) for p in d.glob("*.txt")))
    except Exception:
        return ()


@st.cache_data(show_spinner=False)
def _cached_parse(folder: str, qc_json: str, fmt: str, sig: tuple):
    """解析一次实验并缓存。qc_json 为质控名 dict JSON（空=自动识别）；fmt=样品名格式（空=全局设置）。"""
    qc_names = json.loads(qc_json) if qc_json else None
    return P.parse_experiment(config.EXPERIMENTS_DIR / folder, qc_names=qc_names,
                              sample_pattern=fmt or None)


@st.cache_data(show_spinner=False)
def _std_compare_data(folder: str, sig: tuple):
    """原始文件中全部"非样品名"进样（标准品/空白/洗针等）的 (名, 针号)×物质 峰面积。
    供「标准品对比」任选两针对比响应；不走质控名匹配，任何名字都可作对比对象。
    返回 (长表 DataFrame[sample, injection, analyte, area], [(名, 针号) 候选列表])。
    仅支持 ALLDATA 导出（AREA/RT 宽表无进样针粒度）。
    """
    rf = P.find_raw_files(config.EXPERIMENTS_DIR / folder)
    if not rf.alldata:
        return pd.DataFrame(), []
    raw = P.parse_alldata(rf.alldata)

    def _ion_idx(comp: str) -> int:
        m = re.search(r"-(\d+)$", comp or "")
        return int(m.group(1)) if m else 0
    raw = raw[raw["component"].map(_ion_idx) == config.DEFAULT_ION_PAIR].copy()

    def _is_sample(n: str) -> bool:
        return P.parse_sample_name(n) is not None
    cand = raw[~raw["sample"].map(_is_sample)].copy()
    if len(cand) == 0:
        return pd.DataFrame(), []
    # 针号：同名多针按出现顺序（与 parser 的 injection_idx 口径一致）
    cand["injection"] = cand.groupby(["sample", "component"]).cumcount() + 1
    keep = cand[["sample", "injection", "analyte", "area"]].copy()
    keep = keep[keep["analyte"].astype(str) != ""]
    candidates = sorted(
        {(str(s), int(i)) for s, i in zip(keep["sample"], keep["injection"])},
        key=lambda t: (0 if "std" in t[0].lower() else 1, t[0], t[1]))  # 含 STD 的排前
    return keep, candidates


# ---------- 通用工具 ----------
@lru_cache(maxsize=1)
def _analyte_numeric_cols_cached() -> tuple:
    if not REF_OK:
        return ()
    df = R.get_analyte_basic()
    if df is None or len(df) == 0:
        return ()
    cols = []
    for c in df.columns:
        vals = df[c].dropna()
        if len(vals) == 0:
            continue
        n_float = sum(R._try_float(v) is not None for v in vals)
        if n_float / len(vals) >= 0.6:
            cols.append(c)
    return tuple(cols)


def get_analyte_numeric_cols() -> list:
    """分析物的数值型属性列（自定义看板X轴候选；缓存，信息库刷新时清除）"""
    return list(_analyte_numeric_cols_cached())
    df = R.get_analyte_basic()
    if df is None or len(df) == 0:
        return []
    cols = []
    for c in df.columns:
        vals = df[c].dropna()
        if len(vals) == 0:
            continue
        n_float = sum(R._try_float(v) is not None for v in vals)
        if n_float / len(vals) >= 0.6:
            cols.append(c)
    return cols


def dedup_labels(labels: dict) -> dict:
    """组别名去重：重名时加 (M组号) 后缀，确保热图列名唯一"""
    seen = {}
    out = {}
    for g, lbl in labels.items():
        if lbl in seen:
            seen[lbl] += 1
            out[g] = f"{lbl} (M{g})"
        else:
            seen[lbl] = 1
            out[g] = lbl
    # 二次检查（一个名字可能被多个组用）
    counts = {}
    for g, lbl in out.items():
        counts[lbl] = counts.get(lbl, 0) + 1
    final = {}
    for g, lbl in out.items():
        if counts[lbl] > 1:
            final[g] = f"{lbl} (M{g})"
        else:
            final[g] = lbl
    return final


def render_analyte_sidebar(analyte: str, ph_donor: float = None, ph_acceptor: float = None):
    """侧栏：分析物详情（照片+性质+按pH的LogD/电荷数）"""
    if not REF_OK:
        st.sidebar.info(f"已选：{analyte}"); return
    info = R.get_analyte_info(analyte)
    img = R.get_analyte_photo(analyte)
    with st.sidebar:
        st.markdown(f"### 🔬 {analyte}")
        if img:
            st.image(str(img), width="stretch")
        if info is not None:
            # 用户自定义展示列（旧列名自动映射到新表列；都不存在时回退为表内前几列）
            priority = S.get_display_columns("analyte")
            kv = {}
            for c in priority:
                real = R.resolve_column(c, list(info.index))
                if real and info[real] not in (None, "", "None"):
                    kv[real] = info[real]
            if not kv:
                n = 0
                for c in info.index:
                    v = info[c]
                    if v not in (None, "", "None") and n < 8:
                        kv[c] = v
                        n += 1
            if kv:
                st.markdown("**关键性质**")
                st.dataframe(pd.DataFrame(list(kv.items()), columns=["属性", "值"]),
                             width="stretch", hide_index=True)
            with st.expander("查看全部属性"):
                st.dataframe(pd.DataFrame([info]), width="stretch", hide_index=True)
        else:
            st.info("信息库中未找到该物质")
        # —— 按pH的属性（自动识别 analyte.xlsx 中所有 pH 宽表：LogD/电荷数/溶解度…）——
        ph_props = R.list_ph_properties()
        if ph_props:
            st.markdown("**随pH变化的性质**（" + " / ".join(p["display"] for p in ph_props.values()) + "）")
            ph_default = float(ph_donor) if ph_donor else 7.0
            # key 绑定默认值：切换实验/组时 widget 重建，显示该次实验的供体相pH
            ph_input = st.number_input("pH（默认取该组供体相pH）", value=ph_default, step=0.1,
                                       key=f"ph_{analyte}_{ph_default}", min_value=0.0, max_value=14.0)
            kv_ph = {}
            for key, meta_p in ph_props.items():
                disp = meta_p["display"]
                v = R.get_property_by_ph(analyte, disp, ph_input)
                kv_ph[f"{disp} (pH={ph_input:g})"] = f"{v:.4g}" if v is not None else "N/A"
            if ph_acceptor:
                for key, meta_p in ph_props.items():
                    disp = meta_p["display"]
                    v_a = R.get_property_by_ph(analyte, disp, ph_acceptor)
                    kv_ph[f"{disp} (受体相pH={ph_acceptor:g})"] = f"{v_a:.4g}" if v_a is not None else "N/A"
            st.dataframe(pd.DataFrame(list(kv_ph.items()), columns=["属性", "值"]),
                         width="stretch", hide_index=True)


def render_slm_sidebar(slm_name: str):
    """侧栏：SLM 性质详情"""
    if not REF_OK or not slm_name:
        return
    info = R.get_slm_info(slm_name)
    img = R.get_slm_photo(slm_name)
    with st.sidebar:
        st.markdown(f"### 🧪 SLM: {slm_name}")
        if img:
            st.image(str(img), width="stretch")
        if info is not None:
            row = info.iloc[0]
            priority = S.get_display_columns("slm")
            kv = {c: row[c] for c in priority if c in info.columns and row[c] not in (None, "", "None")}
            if kv:
                st.markdown("**关键性质**")
                st.dataframe(pd.DataFrame(list(kv.items()), columns=["属性", "值"]),
                             width="stretch", hide_index=True)
        else:
            st.info(f"未在信息库中找到 SLM「{slm_name}」")


# ========================================================
# 1. 新增实验
# ========================================================
def page_new_experiment():
    st.title("➕ 新增实验")

    st.subheader("1) 选择实验文件夹")
    exp_root = config.EXPERIMENTS_DIR
    if not exp_root.exists():
        exp_root.mkdir(parents=True, exist_ok=True)

    # —— 方式A：直接拖入数据文件（自动创建实验文件夹）——
    with st.expander("📤 直接拖入数据文件（自动创建实验文件夹）"):
        up_files = st.file_uploader(
            "拖入或选择本次实验的 txt 数据文件（ALLDATA / AREA / RT，可多选）",
            type=["txt"], accept_multiple_files=True, key="ne_upload")
        if up_files:
            default_new = f"{date.today().strftime('%Y%m%d')}-{len(up_files)}F"
            new_name = st.text_input("新实验文件夹名", value=default_new, key="ne_new_folder")
            if st.button("💾 保存到实验文件夹并解析", key="ne_save_upload"):
                target = exp_root / new_name
                target.mkdir(parents=True, exist_ok=True)
                saved = []
                for f in up_files:
                    out = target / f.name
                    out.write_bytes(f.getvalue())
                    saved.append(f.name)
                st.success(f"已保存 {len(saved)} 个文件到 `{target}`：{', '.join(saved)}")
                st.rerun()
        st.caption("保存后该文件夹会出现在下方下拉框中，选择即可继续。")

    # —— 方式B：选择已有实验文件夹 ——
    folders = sorted([p.name for p in exp_root.iterdir() if p.is_dir()])
    if not folders:
        st.info(f"请先拖入数据文件，或在 `{exp_root}` 下创建实验文件夹（如 20260809-12M）并放入 ALLDATA/AREA+RT txt。")
        return
    sel = st.selectbox("实验文件夹", ["（请选择）"] + folders, key="ne_folder")
    if sel == "（请选择）":
        return

    folder = exp_root / sel

    # 自动回填：该实验已录入过时，把数据库中保存的录入信息注入表单默认值
    conn_ne = get_db()
    _erow = conn_ne.execute("SELECT * FROM experiments WHERE folder=?", (sel,)).fetchone()
    if st.session_state.get("_ne_loaded_folder") != sel:
        st.session_state["_ne_loaded_folder"] = sel
        # 多标准品组合编辑器控件键（ne_combo_*）不含文件夹维度：切文件夹时清掉，
        # 防止 A 实验的组合名/角色下拉旧值经控件状态串到 B 实验并存库
        import re as _re_combo
        for _k in [k for k in list(st.session_state.keys()) if _re_combo.match(r"^ne_combo_", str(k))]:
            st.session_state.pop(_k, None)
        if _erow is None:
            # 切到未录入过的文件夹：把表单控件显式写回默认值。
            st.session_state["ne_sample_fmt"] = "跟随全局设置"
            st.session_state.pop("ne_sample_fmt_re", None)
            st.session_state["ne_multi_qc"] = False
            st.session_state.pop("ne_qc_combos", None)
            st.session_state.pop("_ne_quick_state", None)
            st.session_state.pop("_ne_detail_state", None)
            st.session_state.pop("_ne_last_group_mode", None)
            import re as _re_sync
            for _k in [k for k in list(st.session_state.keys())
                       if _re_sync.match(r"^g\d+_(ov_|label$|well$|abw$|note$|tags$|qc_combo$|slm_main$|pH_donor$|pH_acceptor$|extract_time$|extract_voltage$|donor_composition$|donor_buffer_conc$|acceptor_composition$|acceptor_buffer_conc$|donor_additives$|acceptor_additives$|slm_additives$|.*_pct_)", k)]:
                st.session_state.pop(_k, None)
            # （不能只删 key——浏览器 rerun 时会把旧控件值提交回来覆盖删除；
            #   显式赋值与"填入"机制对称，确定性生效）
            st.session_state["ne_date"] = date.today()
            st.session_state["ne_temp"] = str(S.get_setting("DEFAULT_TEMPERATURE") or "")
            st.session_state["ne_hum"] = str(S.get_setting("DEFAULT_HUMIDITY") or "")
            st.session_state["ne_cp"] = str(S.get_setting("DEFAULT_COLUMN_PRESSURE") or "")
            st.session_state["ne_va"] = float(S.get_setting("DEFAULT_VA"))
            st.session_state["ne_vd"] = float(S.get_setting("DEFAULT_VD"))
            st.session_state["ne_purpose"] = ""
            st.session_state["ne_cur_sheets"] = []
            st.session_state["ne_current_sheets"] = []
            st.session_state["ne_exp_type"] = list(EXP_TYPES.keys())[0]
            st.session_state["base_slm"] = []
            st.session_state["base_sa"] = []
            st.session_state["base_da"] = []
            st.session_state["base_aa"] = []
            st.session_state["base_pd"] = float(S.get_setting("DEFAULT_PH_DONOR"))
            st.session_state["base_pa"] = float(S.get_setting("DEFAULT_PH_ACCEPTOR"))
            st.session_state["base_et"] = float(S.get_setting("DEFAULT_EXTRACT_TIME"))
            st.session_state["base_ev"] = float(S.get_setting("DEFAULT_EXTRACT_VOLTAGE"))
            st.session_state["base_dc"] = str(S.get_setting("DEFAULT_DONOR_COMPOSITION") or "甲酸")
            st.session_state["base_dbc"] = str(S.get_setting("DEFAULT_DONOR_BUFFER_CONC") or "")
            st.session_state["base_dac"] = str(S.get_setting("DEFAULT_DONOR_ADDITIVE_CONC") or "")
            st.session_state["base_ac"] = str(S.get_setting("DEFAULT_ACCEPTOR_COMPOSITION") or "甲酸")
            st.session_state["base_abc"] = str(S.get_setting("DEFAULT_ACCEPTOR_BUFFER_CONC") or "")
            st.session_state["base_aac"] = str(S.get_setting("DEFAULT_ACCEPTOR_ADDITIVE_CONC") or "")
            st.session_state["_ne_gc_autofill"] = None
            st.session_state.pop(f"ne_qc_override_{sel}", None)  # 该文件夹可能残留的质控名覆盖
            for _i in range(16):  # 各添加剂含量/浓度输入框（自由单位）一并复位
                st.session_state.pop(f"base_sa_pct_{_i}", None)
                st.session_state.pop(f"base_da_pct_{_i}", None)
                st.session_state.pop(f"base_aa_pct_{_i}", None)
        if _erow:
            st.session_state["ne_date"] = (datetime.strptime(_erow["exp_date"], "%Y%m%d").date()
                                           if _erow["exp_date"] else date.today())
            st.session_state["ne_temp_key"] = _erow["temperature"] or ""
            st.session_state["ne_temp"] = _erow["temperature"] or ""
            st.session_state["ne_hum"] = _erow["humidity"] or ""
            st.session_state["ne_cp"] = _erow["column_pressure"] or ""
            if _erow["Va"]:
                st.session_state["ne_va"] = float(_erow["Va"])
            if _erow["Vd"]:
                st.session_state["ne_vd"] = float(_erow["Vd"])
            st.session_state["ne_purpose"] = _erow["purpose"] or ""
            if _erow["current_sheets"]:
                try:
                    st.session_state["ne_cur_sheets"] = [
                        f"{m['file']} :: {m['sheet']}" for m in json.loads(_erow["current_sheets"])]
                except Exception:
                    pass
            # 实验目的分类回填（selectbox 需带 key；此处在其实例化之前）
            _et_rev = {v["key"]: k for k, v in EXP_TYPES.items()}
            if _erow["exp_type"] in _et_rev:
                st.session_state["ne_exp_type"] = _et_rev[_erow["exp_type"]]
            # 多标准品组合回填
            try:
                _qc_combos_saved = json.loads(_erow["qc_combos"]) if _erow["qc_combos"] else []
            except Exception:
                _qc_combos_saved = []
            st.session_state["ne_multi_qc"] = bool(_qc_combos_saved)
            if _qc_combos_saved:
                st.session_state["ne_qc_combos"] = _qc_combos_saved
            # 各组已保存的组合名 → 详细模式下拉预选（键在控件实例化前写入）
            try:
                _gc_qc_df = M.get_group_conditions(conn_ne, int(_erow["id"]))
                if len(_gc_qc_df):
                    for _, _r_qc in _gc_qc_df[_gc_qc_df["grp"] > 0].iterrows():
                        _cn_qc = _r_qc.get("qc_combo")
                        if _cn_qc:
                            st.session_state[f"g{int(_r_qc['grp'])}_qc_combo"] = str(_cn_qc)
            except Exception:
                pass
            # 样品名称格式回填（该实验保存过的格式；否则跟随全局）
            _sf_saved = _erow["sample_format"] if "sample_format" in _erow.keys() else None
            if _sf_saved:
                st.session_state["ne_sample_fmt"] = next(
                    (_k for _k, _rx in S.SAMPLE_FORMAT_PRESETS.items() if _rx == _sf_saved),
                    "自定义正则")
            else:
                st.session_state["ne_sample_fmt"] = "跟随全局设置"
            # 组级条件（基准/各组覆盖/逐组多选）在解析后回填 → 置待办标记
            st.session_state["_ne_autofill_pending"] = True
            st.session_state["_ne_autofill_show_msg"] = True

    # 删除传错数据的实验文件夹（两步确认防误删）
    with st.expander("🗑 删除该实验文件夹（数据传错时用）"):
        st.warning(f"将永久删除 `{folder}` 及其中的全部文件，且不可恢复（数据库记录不受影响）。")
        confirm_del = st.checkbox("我确认要删除该文件夹", key="ne_del_confirm")
        if confirm_del and st.button("🗑 确认删除", type="primary", key="ne_del_btn"):
            import shutil
            shutil.rmtree(folder, ignore_errors=True)
            st.session_state["_flash_msg"] = f"已删除文件夹 `{sel}`。"
            st.session_state.pop(f"ne_qc_override_{sel}", None)
            st.session_state.pop("_std_check_eid", None)
            st.session_state.pop("_std_check_folder", None)
            st.session_state.get("_excl_res_cache", {}).pop(sel, None)
            st.session_state.get("_excl_res_cache_sig", {}).pop(sel, None)
            st.session_state.get("_width_res_cache", {}).pop(sel, None)
            st.rerun()

    # 用户点过「应用质控名重新解析」后，该文件夹的质控名覆盖会存到这里，
    # 之后每次解析（含缓存命中）都按它来，避免重解析被下一次 rerun 覆盖回去
    _qc_override = st.session_state.get(f"ne_qc_override_{sel}")
    # 多组合模式下，编辑器内容变化即并入解析覆盖（不用等按钮）
    _ne_multi_now = st.session_state.get("ne_multi_qc")
    if _ne_multi_now and st.session_state.get("ne_qc_combos"):
        _qc_override = _combos_union_qc(st.session_state["ne_qc_combos"])

    # —— 样品名称格式（本实验解析用；默认跟随全局设置，已录入实验默认其保存格式）——
    _sf_row = None
    try:
        _sf_row = get_db().execute("SELECT sample_format FROM experiments WHERE folder=?", (sel,)).fetchone()
    except Exception:
        _sf_row = None
    _saved_fmt = (_sf_row["sample_format"] if _sf_row else None) or None
    _preset_items = list(S.SAMPLE_FORMAT_PRESETS.items())
    _fmt_labels = ["跟随全局设置"] + [k for k, _ in _preset_items] + ["自定义正则"]
    if _saved_fmt:
        _def_label = next((k for k, rx in _preset_items if rx == _saved_fmt), "自定义正则")
    else:
        _def_label = "跟随全局设置"
    _fc = st.columns([2, 2])
    def _sync_fmt_example():
        # 格式切换时联动更新预览示例（否则预览框还停留在上一个格式的示例名）
        st.session_state["ne_sample_fmt_test"] = _sample_example_for(
            st.session_state.get("ne_sample_fmt", ""))
    _fmt_label = _fc[0].selectbox("样品名称格式（本实验解析用）", _fmt_labels,
                                  index=_fmt_labels.index(_def_label), key="ne_sample_fmt",
                                  on_change=_sync_fmt_example)
    if _fmt_label == "自定义正则":
        _eff_fmt = _fc[1].text_input("自定义正则（g=组号 ph=相位 r=重复号）",
                                     value=_saved_fmt or S.get_sample_format(),
                                     key="ne_sample_fmt_re").strip() or None
    else:
        _eff_fmt = dict(_preset_items).get(_fmt_label)
    if "ne_sample_fmt_test" not in st.session_state:
        st.session_state["ne_sample_fmt_test"] = _sample_example_for(_fmt_label)
    _fmt_test = _fc[1].text_input("样品名解析预览", key="ne_sample_fmt_test")
    try:
        _pp_prev = P.parse_sample_name(_fmt_test, pattern=_eff_fmt) if _eff_fmt else P.parse_sample_name(_fmt_test)
        if _pp_prev:
            st.caption(f"「{_fmt_test}」→ 组{_pp_prev[0]} / {'受体相' if _pp_prev[1] == 'A' else '供体相'} / 重复{_pp_prev[2]}"
                       + ("（其余名称将作为质控/空白候选）" if _eff_fmt is None else ""))
        else:
            st.caption(f"⚠️「{_fmt_test}」不符合当前格式——将被当作质控/空白候选，请检查格式或预览名。")
    except Exception as _fe:
        st.caption(f"⚠️ 正则无效：{_fe}（将按默认格式解析）")
        _eff_fmt = None

    with st.spinner("解析数据中…"):
        try:
            res = _cached_parse(sel, json.dumps(_qc_override, ensure_ascii=False) if _qc_override else "",
                                _eff_fmt or "",
                                _folder_sig(sel))
        except Exception as e:
            st.error(f"解析失败：{e}"); return
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("物质数", len(res.analytes))
    c2.metric("样品组数", len(res.groups))
    c3.metric("重复数", res.n_replicates)
    c4.metric("数据行数", len(res.df))
    if res.warnings:
        with st.expander(f"⚠️ 解析警告（{len(res.warnings)} 条）"):
            for w in res.warnings[:20]:
                st.write("• " + w)
    st.session_state["cur_res"] = res

    # ---- QC（默认值从用户设置读取）----
    st.subheader("2) 质控样 / 空白样名称")
    st.caption("下拉框列出数据文件中的**全部样品名**（⭐/☆ 只是自动推荐标记，任何样品——包括洗针液等——都可选为标准品/空白）；"
               "默认选中自动识别结果，可在下拉中改选。")

    def _qc_candidates(samples):
        """把数据中疑似标准品/空白的样本排前（供下拉优先展示）"""
        def rank(s):
            sl = s.lower()
            if "std" in sl or "standard" in sl or "标准" in s:
                return 0
            if "blank" in sl or "blk" in sl or "空白" in s:
                return 1
            return 2
        return sorted(samples, key=rank)

    cc = st.columns(4)
    # 选项 = 文件中出现的全部样品名（含洗针/其他命名，不排除任何样品）；
    # ⭐自动推荐只决定排序和默认值，不影响可选范围
    _raw_names = list(getattr(res, "all_sample_names", []) or [])
    if not _raw_names:
        _raw_names = res.df["sample"].unique().tolist()
    all_samples = _qc_candidates(sorted(_raw_names))
    def _fmt_sample(s):
        sl = s.lower()
        if "std" in sl or "standard" in sl or "标准" in s:
            return f"⭐ {s}"
        if "blank" in sl or "blk" in sl or "空白" in s:
            return f"☆ {s}"
        return s
    def _sel(label, default, key):
        idx = all_samples.index(default) if default in all_samples else 0
        return cc[["受体标准品", "供体标准品", "受体空白", "供体空白"].index(label)].selectbox(
            label, options=all_samples, index=idx, key=key, format_func=_fmt_sample)
    # 下拉默认值优先用解析时自动识别的名称（宽松 STD/BLK 匹配），其次用户设置
    qc_names = {
        "std_acceptor": _sel("受体标准品", res.qc_names.get("std_acceptor") or S.get_setting("DEFAULT_STD_ACCEPTOR"), "qa"),
        "std_donor":    _sel("供体标准品", res.qc_names.get("std_donor") or S.get_setting("DEFAULT_STD_DONOR"), "qd"),
        "blk_acceptor": _sel("受体空白", res.qc_names.get("blk_acceptor") or S.get_setting("DEFAULT_BLK_ACCEPTOR"), "ba"),
        "blk_donor":    _sel("供体空白", res.qc_names.get("blk_donor") or S.get_setting("DEFAULT_BLK_DONOR"), "bd"),
    }
    # —— 多标准品组合：不同实验组使用不同的供受体相标准品/空白 ——
    _multi_qc = st.checkbox("实验组使用不同的供受体相质控样（多标准品组合）", key="ne_multi_qc",
                            help="勾选后可定义多套标准品/空白组合并命名；"
                                 "在「4) 萃取条件 → 逐组详细设置」中为每组选择组合，"
                                 "该组的回收率/转移率将基于所选组合计算。")
    qc_combos = []
    if _multi_qc:
        if "ne_qc_combos" not in st.session_state:
            st.session_state["ne_qc_combos"] = [{"name": "默认组合", **qc_names}]
        st.markdown("**标准品组合定义**（每组在下方第 4 步中选择使用哪套组合）")
        qc_combos = _render_qc_combo_editor(all_samples, st.session_state["ne_qc_combos"], "ne")
        st.session_state["ne_qc_combos"] = qc_combos
        # 组合/下拉变化 → 解析覆盖自动同步（下一轮解析保留全部标准品行）
        if _combos_union_qc(qc_combos) != _qc_override:
            st.session_state[f"ne_qc_override_{sel}"] = _combos_union_qc(qc_combos)
            st.rerun()

    # 解析用质控名：多组合模式用各角色名字的并集（解析器保留全部标准品行），
    # 单模式用下拉选择值；两者都会存入 session 覆盖，避免 rerun 丢失
    _qc_for_parse = (_combos_union_qc(qc_combos) if _multi_qc and qc_combos else dict(qc_names))
    if st.button("🔄 应用质控名重新解析"):
        st.session_state[f"ne_qc_override_{sel}"] = _qc_for_parse
        st.rerun()

    # ---- 元数据（日期用日历）----
    st.subheader("3) 实验元数据")
    mc = st.columns(4)
    meta = {}
    # 日历选择日期
    m = re.match(r"^(\d{4})(\d{2})(\d{2})", sel)
    default_d = date(int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else date.today()
    picked = mc[0].date_input("实验日期", value=default_d, key="ne_date")
    meta["exp_date"] = picked.strftime("%Y%m%d")
    meta["temperature"] = mc[1].text_input("温度(℃)", value=str(S.get_setting("DEFAULT_TEMPERATURE") or ""), key="ne_temp")
    meta["humidity"] = mc[2].text_input("湿度(%)", value=str(S.get_setting("DEFAULT_HUMIDITY") or ""), key="ne_hum")
    meta["column_pressure"] = mc[3].text_input("液相平衡柱压", value=str(S.get_setting("DEFAULT_COLUMN_PRESSURE") or ""), key="ne_cp")
    mc2 = st.columns(2)
    meta["Va"] = mc2[0].number_input("受体相体积 Va (μL)", value=float(S.get_setting("DEFAULT_VA")), key="ne_va")
    meta["Vd"] = mc2[1].number_input("供体相体积 Vd (μL)", value=float(S.get_setting("DEFAULT_VD")), key="ne_vd")
    meta["use_sample"] = 1 if st.checkbox("使用样品（启用供体相，计算转移率）", value=True) else 0
    meta["purpose"] = st.text_area("实验说明", value="", key="ne_purpose")

    # ---- 电流记录关联 ----
    st.subheader("3.5) 电流记录关联（可选）")
    st.caption("从 `data/current/` 的 xlsx 中选择本次实验对应的电流 sheet（可多选，如测试萃取时间时多个 sheet）；"
               "也可以直接上传新的电流 xlsx。不选择则按命名规则自动匹配（sheet名=文件夹名）。")
    # 上传新电流 xlsx
    cur_up = st.file_uploader("上传电流 xlsx（保存到 data/current/）", type=["xlsx"], key="ne_cur_upload")
    if cur_up:
        if st.button("💾 保存上传的电流文件", key="ne_cur_save"):
            CUR.CURRENT_DIR.mkdir(parents=True, exist_ok=True)
            (CUR.CURRENT_DIR / cur_up.name).write_bytes(cur_up.getvalue())
            st.success(f"已保存 `{cur_up.name}` 到 data/current/")
            st.rerun()
    # 扫描现有 sheet 供选择
    cur_sheet_list = CUR.list_all_sheets_detailed()
    if cur_sheet_list:
        _opts = [f"{d['file']} :: {d['sheet']}" for d in cur_sheet_list]
        sel_sheets = st.multiselect("选择本次实验的电流 sheet（可多选）", _opts, key="ne_cur_sheets")
        st.session_state["ne_current_sheets"] = [
            {"file": s.split(" :: ")[0], "sheet": s.split(" :: ", 1)[1]} for s in sel_sheets]
        if sel_sheets:
            st.caption(f"已选 {len(sel_sheets)} 个 sheet，将随实验一起保存，实验详情页的电流曲线优先使用。")
    else:
        st.info("`data/current/` 文件夹为空或不存在。可上传 xlsx，或之后把电流文件放进去。")
        st.session_state["ne_current_sheets"] = []

    # ---- 组级萃取条件 ----
    if not res.groups:
        st.warning("⚠️ 当前样品名称格式下没有识别到任何样品组（所有名称都被当作质控/空白候选）。"
                   "请检查上方「样品名称格式」是否与数据的命名方式一致。")
        return
    st.subheader("4) 萃取条件（基准 + 各组覆盖）")
    st.caption("💡 先填一组**基准条件**；各组里**勾选「覆盖」**才会修改该字段，不勾选则继承基准。")

    exp_type_label = st.selectbox("实验目的分类", list(EXP_TYPES.keys()), key="ne_exp_type")
    exp_type = EXP_TYPES[exp_type_label]
    st.info(f"📋 {exp_type['hint']}")

    slm_main_opts, slm_add_opts = ["（无）"], []
    phase_add_opts = []  # 添加剂选项 = SLM表全部物质（主SLM各表+添加剂各表，三处添加剂共用）
    all_reagent_opts = []  # 试剂数据库全名单（主SLM+添加剂，供成分下拉）
    if REF_OK:
        lib = R.get_slm_library()
        slm_main_opts = ["（无）"] + lib.get("main_names", [])
        phase_add_opts = sorted(set(lib.get("main_names", []) + lib.get("additive_names", [])))
        slm_add_opts = list(phase_add_opts)  # SLM膜中添加剂同样可选全部物质
        all_reagent_opts = phase_add_opts
    # 自定义试剂标签：用户曾自由输入、且尚未出现在试剂表中的名称 → 并入各下拉，
    # 显示时带 🔧 标识（存库值仍是裸名）；与试剂表重名的自动"转正"（不再带标识）
    _lib_all_set = set(all_reagent_opts)
    _custom_set = {n for n in S.get_custom_reagents() if n not in _lib_all_set}
    if _custom_set:
        slm_main_opts = slm_main_opts + sorted(_custom_set - set(slm_main_opts))
        phase_add_opts = sorted(set(phase_add_opts) | _custom_set)
        slm_add_opts = list(phase_add_opts)
        all_reagent_opts = sorted(set(all_reagent_opts) | _custom_set)

    # —— 自动填入（续）：基准萃取条件 + 各组设置。必须在对应控件实例化之前写 session，
    #    （控件实例化后再写其 key 会 StreamlitAPIException）；选项过滤防止非法默认值 ——
    _af_pending = st.session_state.pop("_ne_autofill_pending", False)
    if st.session_state.pop("_ne_autofill_show_msg", False):
        st.info("📂 检测到实验已保存的录入信息，表单已自动填入（重新生成时会覆盖保存）。")
    if _af_pending and _erow is None:
        st.session_state.pop("_ne_gc_autofill", None)
    if _af_pending and _erow is not None:
        try:
            _gc_ne = M.get_group_conditions(get_db(), int(_erow["id"]))
        except Exception:
            _gc_ne = None
        if _gc_ne is not None and len(_gc_ne):
            _b_rows = _gc_ne[_gc_ne["grp"] == 0]
            _b = _b_rows.iloc[0] if len(_b_rows) else {}
            _grow_ne = _gc_ne[_gc_ne["grp"] > 0]
            _exp_key_ne = exp_type["key"]
            _slm_opts_ne = [o for o in slm_main_opts if o != "（无）"]

            def _jload_ne(v):
                try:
                    return json.loads(v) if v else []
                except Exception:
                    return []

            def _bare_ne(x):
                return R._strip_ratio(x)

            def _bget(c):
                return _b.get(c) if len(_b) else None

            # ---- 基准条件（grp=0）----
            _slm_opts_ne_set = set(_slm_opts_ne)
            _ph_opts_set = set(phase_add_opts)
            st.session_state["base_slm"] = [v for v in
                                            (_bare_ne(x) for x in str(_bget("slm_main") or "").split("+"))
                                            if v in _slm_opts_ne_set]
            st.session_state["base_sa"] = [v for v in (_bare_ne(x) for x in _jload_ne(_bget("slm_additives")))
                                           if v in _ph_opts_set]
            st.session_state["base_da"] = [v for v in (_bare_ne(x) for x in _jload_ne(_bget("donor_additives")))
                                           if v in _ph_opts_set]
            st.session_state["base_aa"] = [v for v in (_bare_ne(x) for x in _jload_ne(_bget("acceptor_additives")))
                                           if v in _ph_opts_set]
            # 各添加剂的含量/浓度文本（自由单位）回填到对应输入框
            for _i_, _e_ in enumerate(_jload_ne(_bget("donor_additives"))):
                _c_, _p_ = R.split_name_ratio(str(_e_))
                st.session_state[f"base_da_pct_{_i_}"] = _p_ if _p_ is not None else ""
            for _i_, _e_ in enumerate(_jload_ne(_bget("acceptor_additives"))):
                _c_, _p_ = R.split_name_ratio(str(_e_))
                st.session_state[f"base_aa_pct_{_i_}"] = _p_ if _p_ is not None else ""
            for _k, _c in [("base_pd", "pH_donor"), ("base_pa", "pH_acceptor"),
                           ("base_et", "extract_time"), ("base_ev", "extract_voltage")]:
                _v = _bget(_c)
                if _v not in (None, ""):
                    try:
                        st.session_state[_k] = float(_v)
                    except (TypeError, ValueError):
                        pass
            for _k, _c in [("base_dc", "donor_composition"), ("base_dbc", "donor_buffer_conc"),
                           ("base_dac", "donor_additive_conc"), ("base_ac", "acceptor_composition"),
                           ("base_abc", "acceptor_buffer_conc"), ("base_aac", "acceptor_additive_conc")]:
                _v = _bget(_c)
                if _v not in (None, "") and (_k != "base_dc" or str(_v) in all_reagent_opts):
                    st.session_state[_k] = str(_v)

            # ---- 各组设置 ----
            if len(_grow_ne) and res.groups:
                _grp_map = {int(r["grp"]): r for _, r in _grow_ne.iterrows()}
                _edited_rows = {}
                for _i, _g in enumerate(res.groups):
                    _r = _grp_map.get(int(_g))
                    if _r is None:
                        continue
                    _edit = {}
                    if _r.get("group_label") and str(_r["group_label"]) != f"M{int(_g)}":
                        _edit["组别名(留空自动)"] = str(_r["group_label"])  # 纯 M组号=自动回退名，不预填
                    if _r.get("group_note"):
                        _edit["备注(手动输入)"] = str(_r["group_note"])
                    for _cn, _cd in [("供体相pH", "pH_donor"), ("受体相pH", "pH_acceptor"),
                                     ("萃取时间(分钟)", "extract_time"), ("萃取电压(V)", "extract_voltage")]:
                        try:
                            _v = float(_r.get(_cd))
                            _bv = float(_bget(_cd))
                            if abs(_v - _bv) > 1e-9:
                                _edit[_cn] = _v
                        except (TypeError, ValueError):
                            pass
                    _slm_cell = _bare_ne(str(_r.get("slm_main") or ""))
                    if _slm_cell and _slm_cell != _bare_ne(str(_bget("slm_main") or ""))                             and _slm_cell in _slm_opts_ne_set:
                        _edit["主SLM"] = _slm_cell
                    if _exp_key_ne in ("donor", "all"):
                        _dl = [_bare_ne(x) for x in _jload_ne(_r.get("donor_additives"))]
                        _bl = [_bare_ne(x) for x in _jload_ne(_bget("donor_additives"))]
                        if _dl and _dl != _bl and _dl[0] in _ph_opts_set:
                            _edit["供体相添加剂"] = _dl[0]
                    if _exp_key_ne in ("acceptor", "all"):
                        _al = [_bare_ne(x) for x in _jload_ne(_r.get("acceptor_additives"))]
                        _bl = [_bare_ne(x) for x in _jload_ne(_bget("acceptor_additives"))]
                        if _al and _al != _bl and _al[0] in _ph_opts_set:
                            _edit["受体相添加剂"] = _al[0]
                    if _edit:
                        _edited_rows[_i] = _edit
                # data_editor 不允许 session 预置（ValueAssignmentNotAllowed）→
                # 把各组条件交给快速表格的输入数据构建（表格直接显示保存值）
                st.session_state["_ne_gc_autofill"] = (sel, _grow_ne)
                # 逐组多选（快速表格下方 + all 模式展开区）
                for _g in res.groups:
                    _r = _grp_map.get(int(_g))
                    if _r is None:
                        continue
                    st.session_state[f"qa_{_g}_well"] = M.parse_wells(_r.get("well_position") or "")
                    st.session_state[f"qa_{_g}_tags"] = _jload_ne(_r.get("note_tags"))
                    st.session_state[f"qa_{_g}_abw"] = M.parse_wells(_r.get("abnormal_wells") or "")
                    _slm_g = [v for v in (_bare_ne(x) for x in str(_r.get("slm_main") or "").split("+")) if v]
                    _slm_b = [v for v in (_bare_ne(x) for x in str(_bget("slm_main") or "").split("+")) if v]
                    if _slm_g and _slm_g != _slm_b:
                        st.session_state[f"qa_{_g}_slmmain"] = [v for v in _slm_g if v in _slm_opts_ne_set]
                    st.session_state[f"qa_{_g}_slm"] = [v for v in
                                                        (_bare_ne(x) for x in _jload_ne(_r.get("slm_additives")))
                                                        if v in _ph_opts_set]
                    st.session_state[f"qa_{_g}_don"] = [v for v in
                                                        (_bare_ne(x) for x in _jload_ne(_r.get("donor_additives")))
                                                        if v in _ph_opts_set]
                    st.session_state[f"qa_{_g}_acc"] = [v for v in
                                                        (_bare_ne(x) for x in _jload_ne(_r.get("acceptor_additives")))
                                                        if v in _ph_opts_set]
                    # 逐组详细模式控件（覆盖勾选 + 值）
                    if _r.get("group_label") and str(_r["group_label"]) != f"M{int(_g)}":
                        st.session_state[f"g{_g}_label"] = str(_r["group_label"])
                        st.session_state[f"g{_g}_ov_label"] = True  # 同步勾选手动指定，重生成才不丢
                    st.session_state[f"g{_g}_well"] = M.parse_wells(_r.get("well_position") or "")
                    st.session_state[f"g{_g}_abw"] = M.parse_wells(_r.get("abnormal_wells") or "")
                    st.session_state[f"g{_g}_note"] = str(_r.get("group_note") or "")
                    for _f in ["pH_donor", "pH_acceptor", "extract_time", "extract_voltage"]:
                        try:
                            _v = float(_r.get(_f))
                            _bv = float(_bget(_f))
                            if abs(_v - _bv) > 1e-9:
                                st.session_state[f"g{_g}_ov_{_f}"] = True
                                st.session_state[f"g{_g}_{_f}"] = _v
                        except (TypeError, ValueError):
                            pass
                    for _f in ["donor_composition", "donor_buffer_conc", "donor_additive_conc",
                               "acceptor_composition", "acceptor_buffer_conc", "acceptor_additive_conc"]:
                        _v = str(_r.get(_f) or "")
                        _bv = str(_bget(_f) or "")
                        if _v != _bv:
                            st.session_state[f"g{_g}_ov_{_f}"] = True
                            st.session_state[f"g{_g}_{_f}"] = _v
                    for _f in ["slm_additives", "donor_additives", "acceptor_additives"]:
                        _v = [x for x in (_bare_ne(y) for y in _jload_ne(_r.get(_f))) if x in _ph_opts_set]
                        _bv = [_bare_ne(y) for y in _jload_ne(_bget(_f))]
                        if _v != _bv:
                            st.session_state[f"g{_g}_ov_{_f}"] = True
                            st.session_state[f"g{_g}_{_f}"] = _v
                    _slm_v = _bare_ne(str(_r.get("slm_main") or ""))
                    if _slm_v and _slm_v != _bare_ne(str(_bget("slm_main") or "")) and _slm_v in slm_main_opts:
                        st.session_state[f"g{_g}_ov_slm_main"] = True
                        st.session_state[f"g{_g}_slm_main"] = _slm_v

    def _reagent_fmt(n):
        return f"🔧 {n}（自定义）" if n in _custom_set else n

    def _register_custom_reagent(vals):
        """自由输入的新名称立即登记为自定义试剂（之后所有下拉框都能选到）。"""
        for _v in (vals or []):
            if _v and _v not in _lib_all_set:
                S.add_custom_reagent(_v)

    # 基准条件（默认值从用户设置读取）
    st.markdown("**基准萃取条件**")
    bc = st.columns(4)
    # 按逻辑分组录入，便于阅读
    st.markdown("**基准萃取条件**")
    # --- SLM ---
    st.markdown("🔹 **SLM（支撑液膜）**")
    base = {}
    # 主 SLM 多选 + 比例（多选时显示比例输入框，单选可不填）
    base_slm_sel = st.multiselect("主 SLM（可多选组成混合膜，可直接输入自定义名称）",
                                  [o for o in slm_main_opts if o != "（无）"], key="base_slm",
                                  format_func=_reagent_fmt, accept_new_options=True)
    _register_custom_reagent(base_slm_sel)
    base_slm_entries = []
    if len(base_slm_sel) > 1:
        pct_cols = st.columns(len(base_slm_sel))
        for i, nm in enumerate(base_slm_sel):
            pct = pct_cols[i].number_input(f"{nm} 比例(%)", value=100.0 / len(base_slm_sel),
                                           step=1.0, key=f"base_slm_pct_{i}")
            base_slm_entries.append(f"{nm} ({pct:g}%)")
    elif len(base_slm_sel) == 1:
        base_slm_entries.append(base_slm_sel[0])
    base["slm_main"] = base_slm_entries[0] if len(base_slm_entries) == 1 else " + ".join(base_slm_entries)
    base["slm_main_list"] = base_slm_entries
    bc = st.columns(4)
    base["slm_additives"] = bc[0].multiselect("SLM 添加剂（可选SLM表全部物质）", slm_add_opts,
                                               key="base_sa", format_func=_reagent_fmt,
                                               accept_new_options=True) if slm_add_opts else []
    _register_custom_reagent(base["slm_additives"])
    # SLM 添加剂比例输入
    if base["slm_additives"]:
        pct_cols = st.columns(len(base["slm_additives"]))
        entries = []
        for i, nm in enumerate(base["slm_additives"]):
            pct = pct_cols[i].number_input(f"{nm} 含量(%)", value=5.0, step=0.5,
                                           key=f"base_sa_pct_{i}")
            entries.append(f"{nm} ({pct:g}%)")
        base["slm_additives"] = entries
    # --- 供体相 ---
    st.markdown("🔹 **供体相**")
    bcd = st.columns(4)
    base["pH_donor"] = bcd[0].number_input("供体相 pH", value=float(S.get_setting("DEFAULT_PH_DONOR")), step=0.1, key="base_pd")
    base["donor_additives"] = bcd[1].multiselect("供体相添加剂（可选SLM表全部物质）", phase_add_opts,
                                                  key="base_da", format_func=_reagent_fmt,
                                                  accept_new_options=True) if phase_add_opts else []
    _register_custom_reagent(base["donor_additives"])
    if base["donor_additives"]:
        _d_pct = st.columns(len(base["donor_additives"]))
        _d_ent = []
        for i, nm in enumerate(base["donor_additives"]):
            _p = _d_pct[i].text_input(
                f"供体·{nm} 含量/浓度（单位自定）", key=f"base_da_pct_{i}",
                value=str(S.get_setting("DEFAULT_DONOR_ADDITIVE_CONC") or ""),
                placeholder="例：5% 或 10 mM").strip()
            _d_ent.append(f"{nm} ({_p})" if _p else nm)
        base["donor_additives"] = _d_ent
    # 成分：下拉选试剂数据库名单（默认甲酸）
    _comp_opts = all_reagent_opts or [o for o in slm_main_opts if o != "（无）"]
    _comp_default = str(S.get_setting("DEFAULT_DONOR_COMPOSITION") or "甲酸")
    _comp_idx = _comp_opts.index(_comp_default) if _comp_default in _comp_opts else (_comp_opts.index("甲酸") if "甲酸" in _comp_opts else 0)
    base["donor_composition"] = bcd[2].selectbox("供体相成分（下拉选试剂）", _comp_opts, index=_comp_idx, key="base_dc")
    base["donor_buffer_conc"] = bcd[3].text_input("供体相缓冲剂浓度", value=str(S.get_setting("DEFAULT_DONOR_BUFFER_CONC") or ""), key="base_dbc")
    # --- 受体相 ---
    st.markdown("🔹 **受体相**")
    bca = st.columns(4)
    base["pH_acceptor"] = bca[0].number_input("受体相 pH", value=float(S.get_setting("DEFAULT_PH_ACCEPTOR")), step=0.1, key="base_pa")
    base["acceptor_additives"] = bca[1].multiselect("受体相添加剂（可选SLM表全部物质）", phase_add_opts,
                                                    key="base_aa", format_func=_reagent_fmt,
                                                    accept_new_options=True) if phase_add_opts else []
    _register_custom_reagent(base["acceptor_additives"])
    if base["acceptor_additives"]:
        _a_pct = st.columns(len(base["acceptor_additives"]))
        _a_ent = []
        for i, nm in enumerate(base["acceptor_additives"]):
            _p = _a_pct[i].text_input(
                f"受体·{nm} 含量/浓度（单位自定）", key=f"base_aa_pct_{i}",
                value=str(S.get_setting("DEFAULT_ACCEPTOR_ADDITIVE_CONC") or ""),
                placeholder="例：5% 或 10 mM").strip()
            _a_ent.append(f"{nm} ({_p})" if _p else nm)
        base["acceptor_additives"] = _a_ent
    _acomp_default = str(S.get_setting("DEFAULT_ACCEPTOR_COMPOSITION") or "甲酸")
    _acomp_idx = _comp_opts.index(_acomp_default) if _acomp_default in _comp_opts else (_comp_opts.index("甲酸") if "甲酸" in _comp_opts else 0)
    base["acceptor_composition"] = bca[2].selectbox("受体相成分（下拉选试剂）", _comp_opts, index=_acomp_idx, key="base_ac")
    base["acceptor_buffer_conc"] = bca[3].text_input("受体相缓冲剂浓度", value=str(S.get_setting("DEFAULT_ACCEPTOR_BUFFER_CONC") or ""), key="base_abc")
    # --- 萃取 ---
    st.markdown("🔹 **萃取条件**")
    bce = st.columns(2)
    base["extract_time"] = bce[0].number_input("萃取时间(分钟)", value=float(S.get_setting("DEFAULT_EXTRACT_TIME")), step=1.0, key="base_et")
    base["extract_voltage"] = bce[1].number_input("萃取电压(V)", value=float(S.get_setting("DEFAULT_EXTRACT_VOLTAGE")), step=5.0, key="base_ev")

    # 各组覆盖
    res2 = st.session_state["cur_res"]
    group_full = {}
    base_idx = slm_main_opts.index(base["slm_main"]) if base["slm_main"] in slm_main_opts else 0

    group_mode = st.radio(
        "各组条件录入方式",
        ["📋 表格快速设置（推荐）", "🔧 逐组详细设置（含添加剂/备注）"],
        horizontal=True, key="group_mode")
    st.session_state["_group_mode"] = group_mode

    if group_mode == "📋 表格快速设置（推荐）":
        _slm_opts_tbl = [o for o in slm_main_opts if o != "（无）"] or [""]
        # 供体/受体相添加剂列：下拉选（SLM表全部物质 + 自定义登记试剂；"（无）"=清空）
        _add_opts_tbl = ["（无）"] + phase_add_opts
        _add_free_note = "（下拉选内置试剂，或到基准表单输入新试剂以登记 🔧 自定义标签；选「（无）」=清空该组添加剂）"
        st.caption("一张表编辑所有组：**保持基准值/留空 = 继承基准**；改动 = 覆盖该组该字段。"
                   "表格列随「实验目的」变化（测供体相/受体相时自动出现该相的成分/添加剂列，"
                   f"添加剂列{_add_free_note}）；"
                   "SLM膜中添加剂在表格下方逐组点选（下拉选内置名或输入自定义名）。")
        _exp_key = exp_type["key"]
        # 按实验目的构建动态列
        # 各组保存的条件（自动填入：重新打开已录入实验时表格直接显示保存值）
        _af_grow = None
        _af = st.session_state.get("_ne_gc_autofill")
        if _af and _af[0] == sel:
            _af_grow = _af[1]
        _af_map = {}
        if _af_grow is not None and len(_af_grow):
            _af_map = {int(r["grp"]): r for _, r in _af_grow.iterrows()}
        # 双向同步：上一轮是逐组详细设置 → 行值用详细模式的有效设置（含标准品组合）
        _det_state = st.session_state.get("_ne_detail_state") or {}
        _det_gf = _det_state.get("group_full") if _det_state.get("folder") == sel else None
        _det_map = {int(g_): _gc_ for g_, _gc_ in (_det_gf or {}).items()}
        _apply_det = st.session_state.get("_ne_last_group_mode") == "🔧 逐组详细设置（含添加剂/备注）"
        if _apply_det:
            # 丢弃旧的编辑增量，避免其覆盖同步进来的值
            st.session_state.pop("quick_group_tbl", None)
            # 编辑备份一并失效：否则回放会把切走前的旧快速表格值盖掉详细模式同步的值
            st.session_state.pop("_ne_qtbl_edits", None)
        _rows = []
        for g in res2.groups:
            _r_af = _af_map.get(int(g))
            _gc_det = _det_map.get(int(g)) if _apply_det else None
            # 主SLM单元格：基准为空时留空（不得鹿造第一个试剂名；空=继承基准的空）
            _slm_cell0 = base["slm_main"] if base["slm_main"] else ""
            row = {"组": f"M{g}",
                   "组别名(留空自动)": (str(_r_af.get("group_label")) if _r_af is not None and _r_af.get("group_label")
                                        and str(_r_af.get("group_label")) != f"M{int(g)}" else ""),  # 纯 M组号=自动回退名，不视为手动命名
                   "主SLM": _slm_cell0,
                   "供体相pH": float(base["pH_donor"]), "受体相pH": float(base["pH_acceptor"]),
                   "萃取时间(分钟)": float(base["extract_time"]),
                   "萃取电压(V)": float(base["extract_voltage"])}
            if _r_af is not None:
                # 主SLM（剥比例后须在选项内，否则保持基准——SelectboxColumn 值必须合法）
                _slm_cell_af = R._strip_ratio(str(_r_af.get("slm_main") or ""))
                if _slm_cell_af and _slm_cell_af in _slm_opts_tbl:
                    row["主SLM"] = _slm_cell_af
                for _cn, _cd in [("供体相pH", "pH_donor"), ("受体相pH", "pH_acceptor"),
                                 ("萃取时间(分钟)", "extract_time"), ("萃取电压(V)", "extract_voltage")]:
                    try:
                        row[_cn] = float(_r_af.get(_cd))
                    except (TypeError, ValueError):
                        pass
            if _gc_det is not None:
                # 来自逐组详细模式的有效设置（用户刚在该模式中确认过的值）
                if _gc_det.get("slm_main"):
                    _slm_det = R._strip_ratio(str(_gc_det["slm_main"]))
                    if _slm_det in _slm_opts_tbl:
                        row["主SLM"] = _slm_det
                for _cn, _cd in [("供体相pH", "pH_donor"), ("受体相pH", "pH_acceptor"),
                                 ("萃取时间(分钟)", "extract_time"), ("萃取电压(V)", "extract_voltage")]:
                    try:
                        row[_cn] = float(_gc_det.get(_cd))
                    except (TypeError, ValueError):
                        pass
                if _gc_det.get("group_label") and str(_gc_det["group_label"]) != f"M{int(g)}":
                    row["组别名(留空自动)"] = str(_gc_det["group_label"])  # 纯 M组号=自动回退名，不预填
                if _gc_det.get("group_note"):
                    row["备注(手动输入)"] = str(_gc_det["group_note"])
            if _exp_key in ("donor", "all"):
                _da0 = R._strip_ratio(base["donor_additives"][0]) if base.get("donor_additives") else ""
                _da0_content = ""
                if base.get("donor_additives"):
                    _, _da0_content = R.split_name_ratio(base["donor_additives"][0])
                _da_af = ""
                _da_af_content = ""
                if _r_af is not None:
                    _da_entries_af = _as_name_list(_r_af.get("donor_additives"))
                    if _da_entries_af:
                        _da_af, _da_af_content = R.split_name_ratio(_da_entries_af[0])
                        _da_af = R._strip_ratio(_da_af)
                _cell_upd = {
                    "供体相成分": (str(_r_af.get("donor_composition")) if _r_af is not None and _r_af.get("donor_composition") else (base.get("donor_composition") or "甲酸")),
                    "供体相缓冲剂浓度": (str(_r_af.get("donor_buffer_conc")) if _r_af is not None and _r_af.get("donor_buffer_conc") not in (None, "") else str(base.get("donor_buffer_conc") or "")),
                    "供体相添加剂": (_da_af if _da_af in phase_add_opts else (_da0 or "（无）")),
                    "供体相含量": (_da_af_content if _da_af_content else (_da0_content or "")),
                }
                if _gc_det is not None:
                    if _gc_det.get("donor_composition"):
                        _cell_upd["供体相成分"] = str(_gc_det["donor_composition"])
                    if _gc_det.get("donor_buffer_conc") not in (None, ""):
                        _cell_upd["供体相缓冲剂浓度"] = str(_gc_det["donor_buffer_conc"])
                    _det_entries = _as_name_list(_gc_det.get("donor_additives"))
                    if _det_entries:
                        _det_name, _det_content = R.split_name_ratio(_det_entries[0])
                        _det_name = R._strip_ratio(_det_name)
                        _cell_upd["供体相添加剂"] = _det_name if _det_name in phase_add_opts else _det_name
                        _cell_upd["供体相含量"] = _det_content or ""
                row.update(_cell_upd)
            if _exp_key in ("acceptor", "all"):
                _aa0 = R._strip_ratio(base["acceptor_additives"][0]) if base.get("acceptor_additives") else ""
                _aa0_content = ""
                if base.get("acceptor_additives"):
                    _, _aa0_content = R.split_name_ratio(base["acceptor_additives"][0])
                _aa_af = ""
                _aa_af_content = ""
                if _r_af is not None:
                    _aa_entries_af = _as_name_list(_r_af.get("acceptor_additives"))
                    if _aa_entries_af:
                        _aa_af, _aa_af_content = R.split_name_ratio(_aa_entries_af[0])
                        _aa_af = R._strip_ratio(_aa_af)
                _cell_upd_a = {
                    "受体相成分": (str(_r_af.get("acceptor_composition")) if _r_af is not None and _r_af.get("acceptor_composition") else (base.get("acceptor_composition") or "甲酸")),
                    "受体相缓冲剂浓度": (str(_r_af.get("acceptor_buffer_conc")) if _r_af is not None and _r_af.get("acceptor_buffer_conc") not in (None, "") else str(base.get("acceptor_buffer_conc") or "")),
                    "受体相添加剂": (_aa_af if _aa_af in phase_add_opts else (_aa0 or "（无）")),
                    "受体相含量": (_aa_af_content if _aa_af_content else (_aa0_content or "")),
                }
                if _gc_det is not None:
                    if _gc_det.get("acceptor_composition"):
                        _cell_upd_a["受体相成分"] = str(_gc_det["acceptor_composition"])
                    if _gc_det.get("acceptor_buffer_conc") not in (None, ""):
                        _cell_upd_a["受体相缓冲剂浓度"] = str(_gc_det["acceptor_buffer_conc"])
                    _det_entries_a = _as_name_list(_gc_det.get("acceptor_additives"))
                    if _det_entries_a:
                        _det_name_a, _det_content_a = R.split_name_ratio(_det_entries_a[0])
                        _det_name_a = R._strip_ratio(_det_name_a)
                        _cell_upd_a["受体相添加剂"] = _det_name_a if _det_name_a in phase_add_opts else _det_name_a
                        _cell_upd_a["受体相含量"] = _det_content_a or ""
                row.update(_cell_upd_a)
            # 备注与组别名同理：详细模式同步值已在上方写入，不得再被数据库自动填入值覆盖
            if not _apply_det:
                row["备注(手动输入)"] = (str(_r_af.get("group_note")) if _r_af is not None and _r_af.get("group_note") else "")
            _rows.append(row)
        _tbl = pd.DataFrame(_rows).set_index("组")
        # —— 快速表格编辑备份：data_editor 的未提交编辑会在传入数据变化/重渲染时被重置，
        #    导致用户填好的单元格悄悄丢失（20261030 实测）。每次渲染后把与原始值的差异
        #    存入会话，下次构建表格时回放，保证任何重跑都不丢编辑。——
        _tbl_pristine = _tbl.copy()
        # 失效签名：换文件夹/实验目的/组结构变化（重解析可能改变组语义）→ 旧编辑整体作废，
        # 同时丢弃控件增量，防止上一个实验的未提交编辑被套到新实验的表上
        _qbk_sig = (sel, _exp_key, tuple(int(x) for x in res2.groups))
        _qbk = st.session_state.setdefault("_ne_qtbl_edits", {})
        if st.session_state.get("_ne_qtbl_edits_sig") != _qbk_sig:
            _qbk.clear()
            st.session_state.pop("quick_group_tbl", None)
            st.session_state["_ne_qtbl_edits_sig"] = _qbk_sig
        for (_gi_bk, _ci_bk), _v_bk in list(_qbk.items()):
            if _gi_bk in _tbl.index and _ci_bk in _tbl.columns:
                _tbl.loc[_gi_bk, _ci_bk] = _v_bk
        _slm_opts_cell = ([""] + _slm_opts_tbl) if not base["slm_main"] else _slm_opts_tbl
        _colcfg = {
            "主SLM": st.column_config.SelectboxColumn(options=_slm_opts_cell),
            "供体相pH": st.column_config.NumberColumn(format="%.2f"),
            "受体相pH": st.column_config.NumberColumn(format="%.2f"),
            "萃取时间(分钟)": st.column_config.NumberColumn(format="%.1f"),
            "萃取电压(V)": st.column_config.NumberColumn(format="%.1f"),
        }
        if _exp_key in ("donor", "all"):
            _colcfg.update({
                "供体相成分": st.column_config.SelectboxColumn(options=_comp_opts),
                "供体相添加剂": st.column_config.SelectboxColumn(
                    "供体相添加剂", options=_add_opts_tbl,
                    help="下拉选试剂或「（无）」=清空；SLM表外试剂请改用下方展开点选（可自由输入并登记）"),
                "供体相含量": st.column_config.TextColumn(
                    "供体相含量", help="单位自定，如 5% / 10 mM；与添加剂名合并保存"),
            })
        if _exp_key in ("acceptor", "all"):
            _colcfg.update({
                "受体相成分": st.column_config.SelectboxColumn(options=_comp_opts),
                "受体相添加剂": st.column_config.SelectboxColumn(
                    "受体相添加剂", options=_add_opts_tbl,
                    help="下拉选试剂或「（无）」=清空；SLM表外试剂请改用下方展开点选（可自由输入并登记）"),
                "受体相含量": st.column_config.TextColumn(
                    "受体相含量", help="单位自定，如 5% / 10 mM；与添加剂名合并保存"),
            })
        _well_opts = ["（未记录）"] + [f"{r}{c}" for r in "ABCDEFGH" for c in range(1, 13)]
        _edited = st.data_editor(_tbl, key="quick_group_tbl", width="stretch",
                                 num_rows="fixed", column_config=_colcfg)
        # 差异备份（对照未回放备份的原始表）：编辑过→记录，改回原值→移除
        for _gi_bk in _edited.index:
            if _gi_bk not in _tbl_pristine.index:
                continue  # 索引列可编辑：用户改了组名时跳过该行，避免 diff KeyError
            for _ci_bk in _edited.columns:
                _nv = _edited.at[_gi_bk, _ci_bk]
                _pv = _tbl_pristine.at[_gi_bk, _ci_bk]
                _same = bool(pd.isna(_nv) and pd.isna(_pv)) or (_nv == _pv)
                _key_bk = (_gi_bk, _ci_bk)
                if _same:
                    _qbk.pop(_key_bk, None)
                else:
                    _qbk[_key_bk] = None if pd.isna(_nv) else _nv
        # SLM膜中添加剂 + 异常标签：逐组 multiselect（下拉+自定义；供体/受体添加剂在表格列或all模式下多选）
        note_tags_opts = S.get_note_tags()
        st.markdown("**各组 SLM 膜中添加剂 + 异常标签**（添加剂可下拉选内置名或输入自定义名；异常标签如渗漏/烧焦/破裂）")
        with st.expander("展开点选各组 孔位 / SLM添加剂 / 异常标签 / 异常孔位", expanded=len(res2.groups) <= 6):
            quick_add = {"slm_additives": {}, "note_tags": {}, "abnormal_wells": {}, "well_positions": {}}
            _all_wells = [f"{r}{c}" for r in "ABCDEFGH" for c in range(1, 13)]
            head_c = st.columns([0.6, 1, 1, 1, 1, 1])
            head_c[0].markdown("**组**")
            head_c[1].markdown("**孔位**（可多选）")
            head_c[2].markdown("**主SLM组合**")
            head_c[3].markdown("**SLM添加剂**（可多选+比例）")
            head_c[4].markdown("**异常标签**")
            head_c[5].markdown("**异常孔位**（可选）")
            for g in res2.groups:
                row_c = st.columns([0.6, 1, 1, 1, 1, 1])
                row_c[0].markdown(f"**M{g}**")
                quick_add["well_positions"][g] = row_c[1].multiselect(
                    f"M{g} 孔位（可多选，下拉选择）", _all_wells,
                    default=[], key=f"qa_{g}_well",
                    placeholder="下拉选 A1~H12，重复几个选几个")
                # 主SLM组合（多选+比例，默认继承基准）
                _slm_main_opts_g = [o for o in slm_main_opts if o != "（无）"]
                _g_slm_sel = row_c[2].multiselect(
                    f"M{g} 主SLM组合（可多选）", _slm_main_opts_g,
                    default=[x for x in base_slm_sel if x in _slm_main_opts_g], key=f"qa_{g}_slmmain",
                    placeholder="留空=继承基准", format_func=_reagent_fmt,
                    accept_new_options=True)
                _register_custom_reagent(_g_slm_sel)
                _g_slm_entries = []
                if _g_slm_sel and _g_slm_sel != base_slm_sel:
                    if len(_g_slm_sel) > 1:
                        _g_pcts = row_c[2].columns(len(_g_slm_sel))
                        for i_, nm_ in enumerate(_g_slm_sel):
                            _p_ = _g_pcts[i_].number_input(f"M{g}·{nm_} %", value=round(100.0/len(_g_slm_sel), 1),
                                                           step=1.0, key=f"qa_{g}_slmp_{i_}",
                                                           label_visibility="collapsed")
                            _g_slm_entries.append(f"{nm_} ({_p_:g}%)")
                    else:
                        _g_slm_entries = list(_g_slm_sel)
                quick_add.setdefault("slm_main_combo", {})[g] = _g_slm_entries
                # SLM添加剂（多选+比例）
                quick_add["slm_additives"][g] = row_c[3].multiselect(
                    f"M{g} SLM添加剂（可多选，含SLM表全部物质）", slm_add_opts,
                    default=[x for x in (R._strip_ratio(v) for v in (base.get("slm_additives") or [])) if x in slm_add_opts],
                    key=f"qa_{g}_slm", accept_new_options=True, placeholder="可下拉选或输入",
                    format_func=_reagent_fmt)
                _register_custom_reagent(quick_add["slm_additives"][g])
                if quick_add["slm_additives"][g]:
                    _sa_pct = row_c[3].columns(len(quick_add["slm_additives"][g]))
                    _sa_ent = []
                    for i_, nm_ in enumerate(quick_add["slm_additives"][g]):
                        _nm_clean, _p_def = R.split_name_ratio(nm_)
                        if _p_def is None:
                            _p_def = _base_ratio_map(base.get("slm_additives")).get(_nm_clean)
                        try:
                            _p_num = float(str(_p_def).replace("%", "")) if _p_def is not None else 5.0
                        except ValueError:
                            _p_num = 5.0
                        _p_ = _sa_pct[i_].number_input(f"M{g}·{_nm_clean} 含量%",
                                                       value=_p_num,
                                                       step=0.5, key=f"qa_{g}_sap_{i_}", label_visibility="collapsed")
                        _sa_ent.append(f"{_nm_clean} ({_p_:g}%)")
                    quick_add["slm_additives"][g] = _sa_ent
                quick_add["note_tags"][g] = row_c[4].multiselect(
                    f"M{g} 异常标签（可多选）", note_tags_opts,
                    default=[], key=f"qa_{g}_tags", accept_new_options=True,
                    placeholder="渗漏/烧焦/破裂…")
                # 异常孔位选项联动：已填孔位 → 只显示那些孔；未填 → 全部96孔
                _abw_opts = quick_add["well_positions"].get(g) or _all_wells
                quick_add["abnormal_wells"][g] = row_c[5].multiselect(
                    f"M{g} 异常孔位（可选）", _abw_opts,
                    default=[], key=f"qa_{g}_abw",
                    help="只在部分电极棒/孔发生异常时才需要填；空=该组使用的孔位全部视为异常")
        if _exp_key in ("donor", "all"):
            with st.expander("展开点选各组供体相添加剂（可多选+含量，补充表格单选列；all模式用）", expanded=False):
                quick_add["donor_additives"] = {}
                for g in res2.groups:
                    _don_sel = st.multiselect(
                        f"M{g} 供体相添加剂（可多选，含SLM表全部物质）", phase_add_opts,
                        default=[x for x in (R._strip_ratio(v) for v in (base.get("donor_additives") or [])) if x in phase_add_opts],
                        key=f"qa_{g}_don", accept_new_options=True, placeholder="可下拉选或输入",
                        format_func=_reagent_fmt)
                    _register_custom_reagent(_don_sel)
                    if _don_sel:
                        _don_pct = st.columns(len(_don_sel))
                        _don_ent = []
                        for i_, nm_ in enumerate(_don_sel):
                            _nm_clean, _p_def = R.split_name_ratio(nm_)
                            if _p_def is None:
                                _p_def = _base_ratio_map(base.get("donor_additives")).get(_nm_clean)
                            _p_ = _don_pct[i_].text_input(
                                f"M{g}·供体·{_nm_clean} 含量/浓度", key=f"qa_{g}_donp_{i_}",
                                value=str(_p_def) if _p_def else "",
                                placeholder="5% / 10 mM", label_visibility="collapsed").strip()
                            _don_ent.append(f"{_nm_clean} ({_p_})" if _p_ else _nm_clean)
                        quick_add["donor_additives"][g] = _don_ent
                    else:
                        quick_add["donor_additives"][g] = []
        if _exp_key in ("acceptor", "all"):
            with st.expander("展开点选各组受体相添加剂（可多选+含量，补充表格单选列；all模式用）", expanded=False):
                quick_add["acceptor_additives"] = {}
                for g in res2.groups:
                    _acc_sel = st.multiselect(
                        f"M{g} 受体相添加剂（可多选，含SLM表全部物质）", phase_add_opts,
                        default=[x for x in (R._strip_ratio(v) for v in (base.get("acceptor_additives") or [])) if x in phase_add_opts],
                        key=f"qa_{g}_acc", accept_new_options=True, placeholder="可下拉选或输入",
                        format_func=_reagent_fmt)
                    _register_custom_reagent(_acc_sel)
                    if _acc_sel:
                        _acc_pct = st.columns(len(_acc_sel))
                        _acc_ent = []
                        for i_, nm_ in enumerate(_acc_sel):
                            _nm_clean, _p_def = R.split_name_ratio(nm_)
                            if _p_def is None:
                                _p_def = _base_ratio_map(base.get("acceptor_additives")).get(_nm_clean)
                            _p_ = _acc_pct[i_].text_input(
                                f"M{g}·受体·{_nm_clean} 含量/浓度", key=f"qa_{g}_accp_{i_}",
                                value=str(_p_def) if _p_def else "",
                                placeholder="5% / 10 mM", label_visibility="collapsed").strip()
                            _acc_ent.append(f"{_nm_clean} ({_p_})" if _p_ else _nm_clean)
                        quick_add["acceptor_additives"][g] = _acc_ent
                    else:
                        quick_add["acceptor_additives"][g] = []
        for g in res2.groups:
            _row_key = f"M{g}"
            # 「组」索引列可编辑：用户改了组名时 _edited 不含 M{g}，退回构建值防 KeyError 崩页
            r = _edited.loc[_row_key] if _row_key in _edited.index else _tbl.loc[_row_key]
            gc = {}
            # 本组的自动填入行（每组各自取，不得用上一循环残留值）
            _r_af = _af_map.get(int(g))
            # 与基准比较：不同则覆盖
            # 标准品组合：保留逐组详细模式中的选择；无详细状态时保留数据库已存组合
            # （否则老实验在快速表格模式下重新生成会把各组组合清成默认）
            _gc_det_qc = _det_map.get(int(g)) if _apply_det else None
            _af_qc = str(_r_af.get("qc_combo")) if _r_af is not None and _r_af.get("qc_combo") else ""
            gc["qc_combo"] = (str(_gc_det_qc.get("qc_combo")) if _gc_det_qc is not None and _gc_det_qc.get("qc_combo")
                              else _af_qc)
            _lbl_raw = r["组别名(留空自动)"]
            lbl = "" if pd.isna(_lbl_raw) else str(_lbl_raw).strip()
            if lbl in ("None", "nan"):
                lbl = ""
            gc["group_label"] = lbl if lbl else None
            _g_combo = quick_add.get("slm_main_combo", {}).get(g) or []
            if _g_combo:
                gc["slm_main"] = " + ".join(_g_combo) if len(_g_combo) > 1 else _g_combo[0]
            else:
                # NaN/空单元格 = 未改动 → 继承基准（NaN 会存成 NULL 或 "nan" 标签）
                _slm_cell = "" if pd.isna(r["主SLM"]) else str(r["主SLM"]).strip()
                if _slm_cell and _slm_cell != str(base["slm_main"]):
                    gc["slm_main"] = _slm_cell
                else:
                    gc["slm_main"] = base["slm_main"]

            def _num_cell(_v, _base_v):
                try:
                    if pd.isna(_v):
                        return _base_v
                except (TypeError, ValueError):
                    return _base_v
                return float(_v) if float(_v) != float(_base_v) else _base_v

            gc["pH_donor"] = _num_cell(r["供体相pH"], base["pH_donor"])
            gc["pH_acceptor"] = _num_cell(r["受体相pH"], base["pH_acceptor"])
            gc["extract_time"] = _num_cell(r["萃取时间(分钟)"], base["extract_time"])
            gc["extract_voltage"] = _num_cell(r["萃取电压(V)"], base["extract_voltage"])
            gc["slm_additives"] = list(quick_add["slm_additives"].get(g) or [])
            # donor/acceptor 动态列
            if _exp_key in ("donor", "all"):
                gc["donor_composition"] = str(r["供体相成分"]) if str(r["供体相成分"]) != str(base.get("donor_composition")) else base.get("donor_composition")
                gc["donor_buffer_conc"] = str(r["供体相缓冲剂浓度"]) if str(r["供体相缓冲剂浓度"]) != str(base.get("donor_buffer_conc") or "") else base.get("donor_buffer_conc")
                tbl_add = str(r["供体相添加剂"])
                # 优先级：表格单元格改动 > 展开点选改动（基准为空时也生效）> 显式清空 > 继承基准
                _base_list = list(base.get("donor_additives") or [])
                _base_da0 = R._strip_ratio(_base_list[0]) if _base_list else ""
                _, _base_da0_content = R.split_name_ratio(_base_list[0]) if _base_list else ("", None)
                _base_da0_content = str(_base_da0_content or "")
                ms_list = list(quick_add.get("donor_additives", {}).get(g) or [])
                _ms_changed = [R._strip_ratio(x) for x in ms_list] != [R._strip_ratio(x) for x in _base_list]
                _tbl_content = str(r.get("供体相含量") or "").strip()
                if tbl_add not in ("（无）", "", "None") and tbl_add != _base_da0:
                    gc["donor_additives"] = [f"{tbl_add} ({_tbl_content})" if _tbl_content else tbl_add]
                elif (tbl_add == _base_da0 and _tbl_content
                      and _tbl_content != _base_da0_content):
                    # 名称同基准但含量改了 → 覆盖
                    gc["donor_additives"] = [f"{tbl_add} ({_tbl_content})"]
                elif _ms_changed:
                    gc["donor_additives"] = ms_list
                elif tbl_add in ("（无）", "", "None"):
                    gc["donor_additives"] = []
                else:
                    gc["donor_additives"] = ms_list if ms_list else _base_list
            if _exp_key in ("acceptor", "all"):
                gc["acceptor_composition"] = str(r["受体相成分"]) if str(r["受体相成分"]) != str(base.get("acceptor_composition")) else base.get("acceptor_composition")
                gc["acceptor_buffer_conc"] = str(r["受体相缓冲剂浓度"]) if str(r["受体相缓冲剂浓度"]) != str(base.get("acceptor_buffer_conc") or "") else base.get("acceptor_buffer_conc")
                tbl_add_a = str(r["受体相添加剂"])
                # 与供体相同规则：表格改动 > 展开点选改动 > 显式清空 > 继承基准
                _base_list_a = list(base.get("acceptor_additives") or [])
                _base_aa0 = R._strip_ratio(_base_list_a[0]) if _base_list_a else ""
                _, _base_aa0_content = R.split_name_ratio(_base_list_a[0]) if _base_list_a else ("", None)
                _base_aa0_content = str(_base_aa0_content or "")
                ms_list_a = list(quick_add.get("acceptor_additives", {}).get(g) or [])
                _ms_changed_a = [R._strip_ratio(x) for x in ms_list_a] != [R._strip_ratio(x) for x in _base_list_a]
                _tbl_content_a = str(r.get("受体相含量") or "").strip()
                if tbl_add_a not in ("（无）", "", "None") and tbl_add_a != _base_aa0:
                    gc["acceptor_additives"] = [f"{tbl_add_a} ({_tbl_content_a})" if _tbl_content_a else tbl_add_a]
                elif (tbl_add_a == _base_aa0 and _tbl_content_a
                      and _tbl_content_a != _base_aa0_content):
                    gc["acceptor_additives"] = [f"{tbl_add_a} ({_tbl_content_a})"]
                elif _ms_changed_a:
                    gc["acceptor_additives"] = ms_list_a
                elif tbl_add_a in ("（无）", "", "None"):
                    gc["acceptor_additives"] = []
                else:
                    gc["acceptor_additives"] = ms_list_a if ms_list_a else _base_list_a
            # 备注（手动文本 + 异常标签 + 孔位 + 异常孔位）
            gc["group_note"] = str(r["备注(手动输入)"]).strip()
            gc["note_tags"] = list(quick_add["note_tags"].get(g) or [])
            gc["well_position"] = json.dumps(
                list(quick_add.get("well_positions", {}).get(g) or []), ensure_ascii=False)
            gc["abnormal_wells"] = json.dumps(
                list(quick_add.get("abnormal_wells", {}).get(g) or []), ensure_ascii=False)
            # 兜底：非 donor/acceptor 目的时动态列不存在，组级字段必须继承基准（否则存 None）
            for k in ["donor_composition", "donor_buffer_conc", "donor_additive_conc",
                      "donor_additives", "acceptor_composition", "acceptor_buffer_conc",
                      "acceptor_additive_conc", "acceptor_additives"]:
                if k not in gc or gc[k] is None:
                    v = base.get(k)
                    gc[k] = list(v) if isinstance(v, list) else v
            group_full[g] = gc
        st.session_state["_ne_quick_state"] = {"folder": sel, "group_full": group_full}
        st.session_state["_ne_last_group_mode"] = "📋 表格快速设置（推荐）"
    else:
        # 各组覆盖
        st.markdown("**各组条件**（勾选「覆盖」才会修改该字段；组别名默认自动生成）")
        res2 = st.session_state["cur_res"]
        group_full = {}
        base_idx = slm_main_opts.index(base["slm_main"]) if base["slm_main"] in slm_main_opts else 0
        # 双向同步：上一轮是快速表格 → 逐组控件按快速表格的有效设置预置
        _quick_state = st.session_state.get("_ne_quick_state") or {}
        _quick_gf = _quick_state.get("group_full") if _quick_state.get("folder") == sel else None
        _apply_quick = st.session_state.get("_ne_last_group_mode") == "📋 表格快速设置（推荐）"

        for g in res2.groups:
            with st.expander(f"M{g}", expanded=(g <= 3)):
                gc = {}
                # —— 从快速表格同步（仅模式切换后的第一轮；之后不再覆写用户在详细模式的改动）——
                if _apply_quick and _quick_gf and int(g) in _quick_gf:
                    _qgc = _quick_gf[int(g)]
                    for _f_ in ["pH_donor", "pH_acceptor", "extract_time", "extract_voltage"]:
                        try:
                            _v_ = float(_qgc.get(_f_))
                            _bv_ = float(base.get(_f_))
                        except (TypeError, ValueError):
                            continue
                        if abs(_v_ - _bv_) > 1e-9:
                            st.session_state[f"g{g}_ov_{_f_}"] = True
                            st.session_state[f"g{g}_{_f_}"] = _v_
                    for _f_ in ["donor_composition", "donor_buffer_conc",
                                "acceptor_composition", "acceptor_buffer_conc"]:
                        _v_ = str(_qgc.get(_f_) or "")
                        _bv_ = str(base.get(_f_) or "")
                        if _v_ and _v_ != _bv_:
                            st.session_state[f"g{g}_ov_{_f_}"] = True
                            st.session_state[f"g{g}_{_f_}"] = _v_
                    for _f_, _cf_ in [("slm_additives", None), ("donor_additives", "供体相含量"),
                                      ("acceptor_additives", "受体相含量")]:
                        _v_list = [R._strip_ratio(x) for x in _as_name_list(_qgc.get(_f_))]
                        _bv_list = [R._strip_ratio(x) for x in _as_name_list(base.get(_f_))]
                        if _v_list and _v_list != _bv_list:
                            st.session_state[f"g{g}_ov_{_f_}"] = True
                            _seed_list = list(_v_list)
                            # 含量列的文本并入第一个添加剂（保留用户在快速表格填的单位）
                            if _cf_ and _seed_list:
                                _ct = str(_qgc.get(_cf_) or "").strip()
                                if _ct:
                                    _seed_list[0] = f"{_seed_list[0]} ({_ct})"
                            st.session_state[f"g{g}_{_f_}"] = _seed_list
                    if _qgc.get("slm_main") and str(_qgc["slm_main"]) != str(base.get("slm_main")):
                        _slm_v_ = R._strip_ratio(str(_qgc["slm_main"]))
                        if _slm_v_ in slm_main_opts:
                            st.session_state[f"g{g}_ov_slm_main"] = True
                            st.session_state[f"g{g}_slm_main"] = _slm_v_
                    if _qgc.get("group_label"):
                        st.session_state[f"g{g}_ov_label"] = True
                        st.session_state[f"g{g}_label"] = str(_qgc["group_label"])
                    if _qgc.get("well_position"):
                        st.session_state[f"g{g}_well"] = M.parse_wells(_qgc["well_position"])
                    if _qgc.get("abnormal_wells"):
                        st.session_state[f"g{g}_abw"] = M.parse_wells(_qgc["abnormal_wells"])
                    if _qgc.get("note_tags"):
                        st.session_state[f"g{g}_tags"] = list(_qgc["note_tags"])
                    if _qgc.get("group_note"):
                        st.session_state[f"g{g}_note"] = str(_qgc["group_note"])
                    if _qgc.get("qc_combo"):
                        st.session_state[f"g{g}_qc_combo"] = str(_qgc["qc_combo"])

                # 标准品组合（多标准品模式：该组回收率/转移率基于所选组合计算）
                _combos_avail = st.session_state.get("ne_qc_combos") or []
                if st.session_state.get("ne_multi_qc") and _combos_avail:
                    _combo_names = [c.get("name") for c in _combos_avail]
                    gc["qc_combo"] = st.selectbox(
                        f"M{g} 标准品组合（供受体相质控）",
                        ["（默认组合）"] + _combo_names,
                        key=f"g{g}_qc_combo",
                        help="该组的回收率/转移率/三相分布将基于所选组合的标准品与空白计算")
                    if gc["qc_combo"] == "（默认组合）":
                        gc["qc_combo"] = ""  # 存空串（默认），不存字面量
                else:
                    gc["qc_combo"] = ""
                # 组别名：自动生成（只读显示，避免 Streamlit widget state 干扰）；
                # 勾选"手动指定"才出现输入框
                gc_manual_override = st.checkbox(f"M{g} 手动指定组别名", key=f"g{g}_ov_label")
                if gc_manual_override:
                    # 基于基准预填一个默认，用户可改
                    auto_prefill = N.generate_group_label(exp_type["key"],
                        slm_main=base["slm_main"], slm_additives=base["slm_additives"],
                        pH_donor=base["pH_donor"], donor_additives=base["donor_additives"],
                        pH_acceptor=base["pH_acceptor"], acceptor_additives=base["acceptor_additives"],
                        extract_time=base["extract_time"], extract_voltage=base["extract_voltage"],
                        donor_composition=base.get("donor_composition") or "",
                        acceptor_composition=base.get("acceptor_composition") or "")
                    custom_label = st.text_input(f"M{g} 组别名", value=auto_prefill or f"M{g}", key=f"g{g}_label")
                    gc["group_label"] = custom_label
                else:
                    # 自动生成：提交时按实际条件生成（此处占位，提交逻辑处理）
                    gc["group_label"] = None  # None = 自动
                    st.caption("组别名将在提交时按该组实际条件自动生成")

                # 通用：覆盖文本字段助手
                def _ov_text(col, oc_idx, field, label, base_val):
                    if col[oc_idx].checkbox(f"覆盖{label}", key=f"g{g}_ov_{field}"):
                        gc[field] = col[oc_idx].text_input(label, value=str(base_val or ""), key=f"g{g}_{field}")
                    else:
                        gc[field] = base_val
                def _ov_num(col, oc_idx, field, label, base_val, step):
                    if col[oc_idx].checkbox(f"覆盖{label}", key=f"g{g}_ov_{field}"):
                        gc[field] = col[oc_idx].number_input(label, value=float(base_val), step=step, key=f"g{g}_{field}")
                    else:
                        gc[field] = base_val
                def _ov_multi(col, oc_idx, field, label, base_val, opts=None, opts_label=None,
                              with_content=False, unit_free=False):
                    _opts = opts if opts is not None else slm_add_opts
                    if _opts and col[oc_idx].checkbox(f"覆盖{label}", key=f"g{g}_ov_{field}"):
                        gc[field] = col[oc_idx].multiselect(
                            opts_label or label, _opts,
                            default=[x for x in (R._strip_ratio(v) for v in (base_val or []))
                                     if x in _opts or _opts is slm_add_opts and R._strip_ratio(x) in _opts],
                            key=f"g{g}_{field}", accept_new_options=True,
                            format_func=_reagent_fmt)
                        _register_custom_reagent(gc[field])
                        if with_content and gc[field]:
                            # 含量/浓度输入（与快速表格一致）：SLM添加剂=数字%；
                            # 供体/受体相=自由单位文本框
                            _p_cols = col[oc_idx].columns(len(gc[field]))
                            _ent = []
                            for _i_, _nm_ in enumerate(gc[field]):
                                _clean_, _p_def_ = R.split_name_ratio(_nm_)
                                if _p_def_ is None:
                                    _p_def_ = _base_ratio_map(base_val).get(_clean_)
                                if unit_free:
                                    _p_ = _p_cols[_i_].text_input(
                                        f"{_clean_} 含量/浓度", key=f"g{g}_{field}_pct_{_i_}",
                                        value=str(_p_def_) if _p_def_ else "",
                                        placeholder="5% / 10 mM",
                                        label_visibility="collapsed").strip()
                                    _ent.append(f"{_clean_} ({_p_})" if _p_ else _clean_)
                                else:
                                    try:
                                        _p_num = float(str(_p_def_).replace("%", "")) if _p_def_ is not None else 5.0
                                    except ValueError:
                                        _p_num = 5.0
                                    _p_ = _p_cols[_i_].number_input(
                                        f"{_clean_} 含量%", value=_p_num, step=0.5,
                                        key=f"g{g}_{field}_pct_{_i_}", label_visibility="collapsed")
                                    _ent.append(f"{_clean_} ({_p_:g}%)")
                            gc[field] = _ent
                    else:
                        gc[field] = list(base_val)
                def _ov_select(col, oc_idx, field, label, base_val):
                    if col[oc_idx].checkbox(f"覆盖{label}", key=f"g{g}_ov_{field}"):
                        v = col[oc_idx].selectbox(label, slm_main_opts, index=base_idx, key=f"g{g}_{field}",
                                                  format_func=_reagent_fmt)
                        gc[field] = "" if v == "（无）" else v
                    else:
                        gc[field] = base_val

                st.caption("SLM")
                oc = st.columns(4)
                _ov_select(oc, 0, "slm_main", "主SLM", base["slm_main"])
                _ov_multi(oc, 1, "slm_additives", "SLM添加剂", base["slm_additives"],
                          with_content=True, unit_free=False)

                st.caption("供体相")
                ocd = st.columns(4)
                _ov_num(ocd, 0, "pH_donor", "供体相pH", base["pH_donor"], 0.1)
                _ov_multi(ocd, 1, "donor_additives", "供体相添加剂", base["donor_additives"],
                          opts=phase_add_opts, opts_label="供体相添加剂（可选SLM表全部物质）",
                          with_content=True, unit_free=True)
                _ov_text(ocd, 2, "donor_composition", "供体相成分", base["donor_composition"])
                _ov_text(ocd, 3, "donor_buffer_conc", "供体相缓冲剂浓度", base["donor_buffer_conc"])
                ocd2 = st.columns(2)

                st.caption("受体相")
                oca = st.columns(4)
                _ov_num(oca, 0, "pH_acceptor", "受体相pH", base["pH_acceptor"], 0.1)
                _ov_multi(oca, 1, "acceptor_additives", "受体相添加剂", base["acceptor_additives"],
                          opts=phase_add_opts, opts_label="受体相添加剂（可选SLM表全部物质）",
                          with_content=True, unit_free=True)
                _ov_text(oca, 2, "acceptor_composition", "受体相成分", base["acceptor_composition"])
                _ov_text(oca, 3, "acceptor_buffer_conc", "受体相缓冲剂浓度", base["acceptor_buffer_conc"])
                oca2 = st.columns(2)

                st.caption("萃取条件")
                oce = st.columns(2)
                _ov_num(oce, 0, "extract_time", "萃取时间", base["extract_time"], 1.0)
                _ov_num(oce, 1, "extract_voltage", "萃取电压", base["extract_voltage"], 5.0)

                st.caption("备注 / 异常标签 / 孔位（重复用多个孔，孔位可多选）")
                nb_c = st.columns(2)
                _well_list = [f"{r}{c}" for r in "ABCDEFGH" for c in range(1, 13)]
                gc["note_tags"] = nb_c[0].multiselect(
                    f"M{g} 异常标签", S.get_note_tags(),
                    default=[], key=f"g{g}_tags", accept_new_options=True,
                    placeholder="渗漏/烧焦/破裂…")
                gc["well_position"] = nb_c[1].multiselect(
                    f"M{g} 孔位(96孔板,可多选)", _well_list,
                    default=[], key=f"g{g}_well", accept_new_options=True,
                    placeholder="该组使用的孔，如 A1、A2、A3")
                nb2_c = st.columns(2)
                gc["abnormal_wells"] = nb2_c[0].multiselect(
                    f"M{g} 异常孔位(可选)", _well_list,
                    default=[], key=f"g{g}_abw",
                    help="只在部分电极棒/孔异常时填；空=该组全部孔位视为异常")
                gc["group_note"] = nb2_c[1].text_input(f"M{g} 备注(手动)", value="", key=f"g{g}_note")

                group_full[g] = gc
        st.session_state["_ne_detail_state"] = {"folder": sel, "group_full": group_full}
        st.session_state["_ne_last_group_mode"] = "🔧 逐组详细设置（含添加剂/备注）"

    # ---- 生成 ----
    st.subheader("5) 生成")
    if st.button("✅ 生成 Excel + 写入数据库", type="primary"):
        conn = get_db()
        # 质控名以下拉当前值为准：与解析覆盖不一致时重解析，保证指标/Excel 与保存的 qc_names 一致
        if dict(qc_names) != (_qc_override or {}):
            st.session_state[f"ne_qc_override_{sel}"] = dict(qc_names)
            res3 = _cached_parse(sel, json.dumps(_qc_for_parse, ensure_ascii=False),
                                 _eff_fmt or "", _folder_sig(sel))
        else:
            res3 = st.session_state["cur_res"]
        # 重新生成已有实验时套用已保存的剔除记录（样本+标准品针），避免被悄悄还原
        _pre_row = conn.execute("SELECT id FROM experiments WHERE folder=?", (sel,)).fetchone()
        _saved_excl_re = M.get_excluded_samples(conn, int(_pre_row["id"])) if _pre_row else set()
        _saved_std_re = M.get_excluded_std(conn, int(_pre_row["id"])) if _pre_row else {}
        if _saved_std_re:
            _drop_re = set().union(*_saved_std_re.values())
            if _drop_re:
                # 用副本承载过滤结果：不得原地改 cur_res——步骤6 需要完整标准品针列表
                # （预勾选已剔除针供取消），原地过滤会让旧剔除在确认时被误清
                import dataclasses as _dc
                _mask_re = res3.df.apply(
                    lambda r_: (str(r_["sample"]), int(r_["injection_idx"])) in _drop_re, axis=1)
                res3 = _dc.replace(res3, df=res3.df[~_mask_re].copy())
        # 多标准品组合：各组 QC 配置（提交时从组条件构建；解析已保留全部标准品行）
        _combos_for_calc = qc_combos if (_multi_qc and qc_combos) else []
        _gqc_submit = {}
        for _g_, _gc_ in group_full.items():
            _cn_ = _gc_.get("qc_combo")
            if _cn_:
                _c_ = next((c for c in _combos_for_calc if c.get("name") == _cn_), None)
                if _c_:
                    _gqc_submit[int(_g_)] = {k: _c_.get(k) for k, _ in QC_ROLES}
        _default_qc_submit = ({k: (v[0] if isinstance(v, list) else v) for k, v in _qc_for_parse.items()}
                              if isinstance(_qc_for_parse.get("std_acceptor"), list) else dict(qc_names))
        metrics = C.compute_metrics(res3.df, res3.analytes, res3.groups,
                                    meta["Va"], meta["Vd"], use_sample=bool(meta["use_sample"]),
                                    excluded_samples=_saved_excl_re,
                                    group_qc=_gqc_submit, default_qc=_default_qc_submit)
        data_xlsx = folder / f"{meta['exp_date'] or res3.file_date}-DATA.xlsx"
        exp_data = {
            "folder": sel, "exp_date": meta["exp_date"],
            "temperature": meta["temperature"], "humidity": meta["humidity"],
            "column_pressure": meta["column_pressure"],
            "Va": meta["Va"], "Vd": meta["Vd"],
            "use_sample": meta["use_sample"], "qc_names": json.dumps(_qc_for_parse, ensure_ascii=False),
            "exp_type": exp_type["key"], "purpose": meta["purpose"],
            "sample_format": _eff_fmt,
            "qc_combos": json.dumps(qc_combos, ensure_ascii=False) if (_multi_qc and qc_combos) else None,
            "current_sheets": json.dumps(st.session_state.get("ne_current_sheets") or [], ensure_ascii=False),
            "n_groups": len(res3.groups), "n_replicates": res3.n_replicates,
            "data_xlsx": str(data_xlsx),
        }
        eid = M.upsert_experiment(conn, exp_data)
        M.save_metrics(conn, eid, meta["exp_date"], metrics, res3.analytes, res3.groups)
        M.save_qc_summary(conn, eid, meta["exp_date"], res3, res3.analytes,
                          default_qc=_default_qc_submit)
        M.save_rt_records(conn, eid, meta["exp_date"], res3)
        def _ser(v):
            return json.dumps(v, ensure_ascii=False) if isinstance(v, list) else v
        base_serial = {**base, "group_label": "基准条件"}
        for k in ["slm_additives", "donor_additives", "acceptor_additives", "note_tags",
                  "well_position", "abnormal_wells"]:
            base_serial[k] = _ser(base_serial.get(k, []))
        gc_serial = {}
        _fallback_groups = []  # 自动命名退回 M组号 的组（缺 SLM 信息），生成后提示
        for g, c in group_full.items():
            gc_serial[g] = {**c}
            for k in ["slm_additives", "donor_additives", "acceptor_additives", "note_tags",
                      "well_position", "abnormal_wells"]:
                gc_serial[g][k] = _ser(gc_serial[g].get(k, []))
            # 组别名：None=自动 → 按该组实际条件生成；否则用户手动指定
            cur_label = c.get("group_label")
            if cur_label is None or str(cur_label).strip() == "":
                regen = N.generate_group_label(exp_type["key"],
                    slm_main=gc_serial[g].get("slm_main"), slm_additives=c.get("slm_additives"),
                    pH_donor=c.get("pH_donor"), donor_additives=c.get("donor_additives"),
                    pH_acceptor=c.get("pH_acceptor"), acceptor_additives=c.get("acceptor_additives"),
                    extract_time=c.get("extract_time"), extract_voltage=c.get("extract_voltage"),
                    donor_composition=c.get("donor_composition") or "",
                    acceptor_composition=c.get("acceptor_composition") or "")
                gc_serial[g]["group_label"] = regen or f"M{g}"
                if not regen:
                    _fallback_groups.append(g)
            else:
                gc_serial[g]["group_label"] = cur_label
        M.save_group_conditions(conn, eid, base_serial, gc_serial)
        # 构建电压表数据：把选中的电流 sheet 数据同步写入 DATA.xlsx（全部选中记录都写入）
        voltage_df = build_voltage_df(st.session_state.get("ne_current_sheets") or [])
        EB.build_excel(data_xlsx, res3, meta, meta["Va"], meta["Vd"], bool(meta["use_sample"]),
                       voltage_df=voltage_df,
                       gcond=M.get_group_conditions(conn, eid),
                       excluded_samples=_saved_excl_re,
                       qc_by_group=_gqc_submit, default_qc=_default_qc_submit)
        n_pts = len(voltage_df) if voltage_df is not None else 0
        # 生成成功（含 Excel 落盘）后才清编辑备份：若 build_excel 抛异常（如 xlsx 被占用），
        # 用户原地重试时表格编辑仍在，不会被迫重填
        st.session_state.pop("_ne_qtbl_edits", None)
        st.session_state.pop("_ne_qtbl_edits_sig", None)
        st.success(f"✅ 完成！Excel：`{data_xlsx}`  实验ID：{eid}"
                   + (f"（电压表已写入 {n_pts} 点电流数据；组条件已同步）" if n_pts else "（组条件已同步）"))
        if _fallback_groups and exp_type.get("key") in N.EXP_TYPE_PREFIX:
            st.warning(
                "⚠️ 组 " + "、".join(f"M{g}" for g in _fallback_groups)
                + " 没有可用的 SLM/膜组成信息，组别名只能用 M组号。"
                "请在「基准条件」或各组设置里填好 SLM 后重新生成本文件夹（组别名会自动补全），"
                "或到「实验详情→修改/删除」手动填写组别名。")
        # 标记进入标准品剔除步骤
        st.session_state["_std_check_eid"] = eid
        st.session_state["_std_check_folder"] = sel

    # —— 6) 标准品剔除步骤（生成后；仅对刚生成的那个文件夹显示）——
    if st.session_state.get("_std_check_folder") != sel:
        st.session_state.pop("_std_check_eid", None)
    std_eid = st.session_state.get("_std_check_eid")
    if std_eid is not None:
        st.subheader("6) 标准品质量检查（可选）")
        conn = get_db()
        qc_df = pd.read_sql("SELECT * FROM qc_summary WHERE experiment_id=?", conn, params=(std_eid,))
        res_chk = st.session_state.get("cur_res")
        exp_row = conn.execute("SELECT data_xlsx, Va, Vd, use_sample, exp_date, qc_names, current_sheets, sample_format, qc_combos, temperature, humidity, column_pressure, purpose FROM experiments WHERE id=?",
                               (std_eid,)).fetchone()
        if len(qc_df) and res_chk is not None and exp_row:
            def _hl_rsd(v):
                if pd.isna(v): return ""
                try:
                    return "background-color: #ffc7ce; color: #9C0006" if float(v) > 15.0 else ""
                except (ValueError, TypeError):
                    return ""

            # 标准品剔除：受体相 / 供体相 两套矩阵分别勾选（支持两相不同标准品环境）
            _std_link = st.checkbox(
                "同一标准品名在两相都使用时联动剔除（勾选 A-STD 第 i 针，另一相的 A-STD 第 i 针同剔；不同名标准品互不影响）",
                value=True, key="std_link_new",
                help="两相用同一标准品系列（同名或同批次对应针）时勾选；只需在一处勾选，两相同号针一起剔除。")
            _needles_by_phase = {}
            _std_names_by_phase = {}
            # 已保存的标准品针剔除（预勾选，取消勾选即恢复；否则二次保存会清掉旧剔除）
            _saved_std_step6 = M.get_excluded_std(conn, std_eid)
            _saved_needles_step6 = {f"{s} #{i}" for pairs in _saved_std_step6.values() for s, i in pairs}
            for phase_cat, phase_name in [("std_acceptor", "受体相标准品"), ("std_donor", "供体相标准品")]:
                _std_names_by_phase[phase_cat] = sorted(
                    res_chk.df[res_chk.df["category"] == phase_cat]["sample"].unique())
                if not _std_names_by_phase[phase_cat]:
                    st.caption(f"（{phase_name}：数据中无对应进样行）")
                    continue
                st.markdown(f"**{phase_name}**（按标准品名分节剔除；剔除某针后，所有使用该标准品的组同时生效）")
                _needles_by_phase.setdefault(phase_cat, [])
                # —— 每个标准品名独立一节：独立矩阵/独立偏离度推荐（不同溶液峰面积不可混算）——
                for _nm in _std_names_by_phase[phase_cat]:
                    _df_nm = res_chk.df[(res_chk.df["category"] == phase_cat)
                                        & (res_chk.df["sample"] == _nm)]
                    std_mat_p = C.get_std_injection_matrix(_df_nm, phase_cat, res_chk.analytes)
                    if not len(std_mat_p):
                        continue
                    with st.expander(f"📌 {_nm}（{len(std_mat_p)} 针）",
                                     expanded=(_nm == _std_names_by_phase[phase_cat][0])):
                        row_sum_p = std_mat_p.sum(axis=1)
                        mean_p = row_sum_p.mean()
                        dev_p = (row_sum_p - mean_p).abs() / (abs(mean_p) if mean_p else 1)
                        worst_p = dev_p.idxmax()
                        st.warning(f"🔴 推荐剔除（偏离最大）：**{worst_p}**（偏离度 {dev_p.max():.3f}）")
                        # 合并表：进样针（索引，冻结）+ 剔除复选框（冻结）+ 峰面积总和/偏离度/各物质（只读）
                        _m_p = std_mat_p.copy()
                        _m_p.insert(0, "峰面积总和", row_sum_p)
                        _m_p.insert(1, "偏离度", dev_p.round(3))
                        _m_p.insert(0, "剔除", [n in _saved_needles_step6 for n in std_mat_p.index])
                        _num_cols_p = [c for c in _m_p.columns if c != "剔除"]
                        _ed_p = st.data_editor(
                            _m_p,
                            key=f"std_ed3_{phase_cat}_{_nm}", hide_index=False,
                            column_config={
                                "剔除": st.column_config.CheckboxColumn("剔除", default=False, pinned=True),
                                "峰面积总和": st.column_config.NumberColumn(format="%.3e"),
                                "偏离度": st.column_config.NumberColumn(format="%.3f"),
                                **{c: st.column_config.NumberColumn(format="%.3e") for c in std_mat_p.columns}},
                            disabled=_num_cols_p)
                        _needles_by_phase[phase_cat].extend(
                            n for n in _ed_p.index[_ed_p["剔除"]].tolist() if n in std_mat_p.index)  # 防误改索引文字

            # 合并两相剔除对（同系列联动：按针号对应扩展到另一相）
            _pairs_by_phase_raw = {}
            for phase_cat, needles in _needles_by_phase.items():
                _pp = []
                for needle in needles:
                    if isinstance(needle, str) and " #" in needle:
                        sname, num = needle.rsplit(" #", 1)
                        _pp.append((sname, int(num)))
                _pairs_by_phase_raw[phase_cat] = _pp
            _pairs_by_phase = _expand_std_link(_pairs_by_phase_raw, _std_names_by_phase) if _std_link else _pairs_by_phase_raw
            excl_pairs = [(s, i) for pairs in _pairs_by_phase.values() for s, i in pairs]
            excl_needles = [f"{s_} #{n_}" for s_, n_ in excl_pairs]

            # —— 实时 RSD 预览：按标准品名分别显示剔除前/后 RSD（不再跨溶液混算）——
            st.markdown("**各物质·各标准品 RSD（勾选剔除后实时更新）**（红色=RSD>15%）")
            _rsd_prev_frames = []
            for phase_cat, phase_label in [("std_acceptor", "受体"), ("std_donor", "供体")]:
                for _nm in _std_names_by_phase.get(phase_cat, []):
                    _excl_nm = {(s, i) for (s, i) in excl_pairs if s == _nm}
                    _df_b = res_chk.df[(res_chk.df["category"] == phase_cat)
                                       & (res_chk.df["sample"] == _nm)]
                    if not len(_df_b):
                        continue
                    _df_a = _df_b[~_df_b.apply(
                        lambda r: (r["sample"], int(r["injection_idx"])) in _excl_nm, axis=1)]
                    _col_b, _col_a = f"{_nm} 剔除前RSD%", f"{_nm} 剔除后RSD%"
                    _rows_pn = []
                    for a in res_chk.analytes:
                        _, r_b, _ = C.aggregate_qc(_df_b, phase_cat, a)
                        _, r_a, _ = C.aggregate_qc(_df_a, phase_cat, a)
                        _rows_pn.append({"物质": a,
                                         _col_b: round(r_b, 2) if r_b is not None else None,
                                         _col_a: round(r_a, 2) if r_a is not None else None})
                    _rdf = pd.DataFrame(_rows_pn)
                    _rdf[f"{_nm} 变化"] = (_rdf[_col_a] - _rdf[_col_b]).round(2)
                    _rsd_prev_frames.append((f"{phase_label}相 · {_nm}", _rdf))
            for _sec, _rdf in _rsd_prev_frames:
                with st.expander(f"{_sec}（{len(_rdf)} 个物质）",
                                 expanded=len(_rsd_prev_frames) <= 2):
                    _cols_rsd = [c for c in _rdf.columns if c != "物质"]
                    _col_after = [c for c in _rdf.columns if "剔除后RSD%" in c][0]
                    # 实时统计置于 RSD 表之前：勾选即可见总数变化
                    _n_over = int((_rdf[_col_after] > 15).sum())
                    _rsd_sum = float(_rdf[_col_after].sum())
                    st.caption(f"📊 实时统计：剔除后 RSD>15% 的物质 **{_n_over}** 个 | "
                               f"所有物质剔除后 RSD% 累加 **{_rsd_sum:.1f}**")
                    st.dataframe(_rdf.style.map(_hl_rsd, subset=_cols_rsd).format(
                        "{:.2f}", na_rep="-", subset=_cols_rsd),
                        width="stretch", hide_index=True)

            if excl_needles and st.button("💾 确认剔除并保存到数据库 + DATA.xlsx", type="primary", key="reapply_std"):
                folder_path = config.EXPERIMENTS_DIR / st.session_state["_std_check_folder"]
                try:
                    qc_names_dic = json.loads(exp_row["qc_names"])
                except Exception:
                    qc_names_dic = {}
                try:
                    _combos_s6 = json.loads(exp_row["qc_combos"]) if exp_row["qc_combos"] else []
                except Exception:
                    _combos_s6 = []
                if _combos_s6:
                    # 多标准品组合：解析需保留全部组合的标准品/空白行
                    qc_names_dic = _combos_union_qc(_combos_s6)
                res_new = _cached_parse(st.session_state["_std_check_folder"],
                                        json.dumps(qc_names_dic, ensure_ascii=False) if qc_names_dic else "",
                                        (exp_row["sample_format"] if exp_row else None) or "",
                                        _folder_sig(st.session_state["_std_check_folder"]))
                drop_mask = res_new.df.apply(
                    lambda r: (r["sample"], int(r["injection_idx"])) in excl_pairs, axis=1)
                res_new.df = res_new.df[~drop_mask].copy()
                va = float(exp_row["Va"]); vd = float(exp_row["Vd"]); us = bool(exp_row["use_sample"])
                # 保留此前已保存的样本剔除，避免重算时把已剔除的重复样本加回来
                _saved_sample_excl = M.get_excluded_samples(conn, std_eid)
                _gqc_s6, _dqc_s6, _ = _build_group_qc_from_db(conn, std_eid)
                m_new = C.compute_metrics(res_new.df, res_new.analytes, res_new.groups, va, vd,
                                          use_sample=us, excluded_samples=_saved_sample_excl,
                                          group_qc=_gqc_s6, default_qc=_dqc_s6)
                conn.execute("DELETE FROM metrics WHERE experiment_id=?", (std_eid,))
                conn.execute("DELETE FROM qc_summary WHERE experiment_id=?", (std_eid,))
                conn.commit()
                M.save_metrics(conn, std_eid, exp_row["exp_date"], m_new, res_new.analytes, res_new.groups)
                M.save_qc_summary(conn, std_eid, exp_row["exp_date"], res_new, res_new.analytes,
                                  default_qc=_dqc_s6)
                # 持久化标准品针剔除（按相分开、已含同系列联动扩展）。
                # 与旧记录合并：本次矩阵中出现并勾选=保留；出现未勾选=恢复；矩阵未出现的旧针=保留
                _final_std = {}
                for _pc_ in ("std_acceptor", "std_donor"):
                    _new_ = set(_pairs_by_phase.get(_pc_) or [])
                    _old_ = _saved_std_step6.get(_pc_) or set()
                    _seen_ = {(s, i) for s, i in _old_ if s in _std_names_by_phase.get(_pc_, set())}
                    _final_std[_pc_] = list(_new_ | (_old_ - _seen_))
                M.set_excluded_std(conn, std_eid,
                                   std_acceptor=_final_std.get("std_acceptor"),
                                   std_donor=_final_std.get("std_donor"))
                data_xlsx_path = Path(exp_row["data_xlsx"]) if exp_row["data_xlsx"] else None
                if data_xlsx_path:
                    try:
                        _meta_s6 = {"exp_date": exp_row["exp_date"],
                                    "temperature": exp_row["temperature"], "humidity": exp_row["humidity"],
                                    "column_pressure": exp_row["column_pressure"], "purpose": exp_row["purpose"]}
                        EB.build_excel(data_xlsx_path, res_new, _meta_s6,
                                       va, vd, us, excluded_samples=_saved_sample_excl,
                                       gcond=M.get_group_conditions(conn, std_eid),
                                       qc_by_group=_gqc_s6, default_qc=_dqc_s6,
                                       voltage_df=build_voltage_df(json.loads(exp_row["current_sheets"]) if exp_row["current_sheets"] else []))
                    except Exception as _ebe:
                        st.session_state["_std_saved_msg"] = (
                            f"⚠️ 数据库已保存，但 DATA.xlsx 生成失败：{_ebe}")
                # 按联动扩展后的实际剔除对计数（同系列/同名标准品时勾任一相，两相同号针同剔）
                n_acc = len(_pairs_by_phase.get("std_acceptor") or [])
                n_don = len(_pairs_by_phase.get("std_donor") or [])
                st.session_state["_std_saved_msg"] = (
                    f"✅ 已剔除受体相标准品 {n_acc} 针、供体相标准品 {n_don} 针，**已保存到数据库**"
                    f"（均值/RSD/回收率已重算），**已重新生成 DATA.xlsx**："
                    f"`{Path(exp_row['data_xlsx']).name if exp_row['data_xlsx'] else '（无文件）'}`。")
                st.rerun()

            # rerun 后显示保存成功提示（一次性：显示后即清除，再次交互不再出现）
            if st.session_state.get("_std_saved_msg"):
                st.success(st.session_state["_std_saved_msg"])
                del st.session_state["_std_saved_msg"]
            # RSD 检查提示（基于 qc_df 最新值）
            n_high = int((qc_df["std_acceptor_rsd"].fillna(0) > config.RSD_THRESHOLD).sum())
            if n_high == 0:
                st.success("✅ 所用物质标准品 RSD 均 ≤ 15%，数据质量良好，无需剔除。")
            elif n_high > 0:
                st.warning(f"⚠️ 有 {n_high} 个物质的标准品 RSD > 15%（见上方表格，红色标记），建议剔除异常进样针。")


# ========================================================
# 2. 实验列表（热图 + 剔除 + 修改）
# ========================================================
def exps_w_analytes(conn, folder: str) -> list:
    """某实验文件夹涉及的分析物列表（解析缓存优先）"""
    cache = st.session_state.setdefault("_excl_res_cache", {})
    res = cache.get(folder)
    if res is None:
        try:
            res = P.parse_experiment(config.EXPERIMENTS_DIR / folder)
            cache[folder] = res
        except Exception:
            return []
    return sorted(res.df["analyte"].unique())


def _page_reagents(conn, folder: str) -> list:
    """某实验文件夹的试剂名集合（主SLM + 各类添加剂）"""
    erow = conn.execute("SELECT id FROM experiments WHERE folder=?", (folder,)).fetchone()
    if not erow:
        return []
    gcond = M.get_group_conditions(conn, int(erow["id"]))
    out = set()
    for _, gr in gcond.iterrows():
        if gr.get("slm_main"):
            out.add(gr["slm_main"])
        for c in ["slm_additives", "donor_additives", "acceptor_additives"]:
            try:
                out.update(x for x in json.loads(gr[c]) if x)
            except Exception:
                pass
    return sorted(out)


@st.cache_data(show_spinner=False)
def _global_detail_options():
    """全库一致的分析物/试剂选项（三页共用同一选项列表，侧栏选择跨页保留）"""
    conn = M.get_conn()
    mdf = pd.read_sql("SELECT DISTINCT analyte FROM metrics", conn)
    gcond = pd.read_sql(
        "SELECT slm_main, slm_additives, donor_additives, acceptor_additives FROM group_conditions", conn)
    analytes = sorted(mdf["analyte"].dropna().unique())
    regs = set()
    for _, r in gcond.iterrows():
        if r["slm_main"]:
            regs.add(r["slm_main"])
        for c in ["slm_additives", "donor_additives", "acceptor_additives"]:
            try:
                regs.update(x for x in json.loads(r[c]) if x)
            except Exception:
                pass
    return analytes, sorted(regs)


def render_detail_sidebar(ph_lookup=None):
    """三页共用的侧栏详情查询，选择跨功能页保留。
    实现：widget 值通过 on_change 回调同步到普通 session_state 变量
    （widget 未渲染时其 key 会被 Streamlit 清理，普通变量不受影响）。
    ph_lookup: 可选函数 物质名→供体相pH（各页传入自己的取值逻辑）。
    """
    if "persist_detail_analyte" not in st.session_state:
        st.session_state["persist_detail_analyte"] = "（不查看）"
    if "persist_detail_slms" not in st.session_state:
        st.session_state["persist_detail_slms"] = []

    def _sync_detail():
        st.session_state["persist_detail_analyte"] = st.session_state.get("w_detail_analyte", "（不查看）")
        st.session_state["persist_detail_slms"] = st.session_state.get("w_detail_slms", [])

    analytes_all, reagents_all = _global_detail_options()
    opts_a = ["（不查看）"] + analytes_all
    cur_a = st.session_state["persist_detail_analyte"]
    if cur_a not in opts_a:
        cur_a = "（不查看）"
    cur_r = [x for x in st.session_state["persist_detail_slms"] if x in reagents_all]

    with st.sidebar:
        st.markdown("---")
        st.markdown("### 📋 详情查看")
        st.selectbox("物质详情", opts_a, index=opts_a.index(cur_a),
                     key="w_detail_analyte", on_change=_sync_detail)
        st.multiselect("SLM/添加剂详情（可多选）", reagents_all, default=cur_r,
                       key="w_detail_slms", on_change=_sync_detail)
        sel_a = st.session_state["persist_detail_analyte"]
        sel_r = st.session_state["persist_detail_slms"]
        if sel_a != "（不查看）":
            ph = ph_lookup(sel_a) if ph_lookup else None
            render_analyte_sidebar(sel_a, ph_donor=ph)
        for r in sel_r:
            _show_slm_card_sidebar(r)


def _show_slm_card_sidebar(name: str):
    """侧栏竖版 SLM 卡片（图在上、性质表在下）"""
    with st.sidebar:
        st.markdown(f"### 🧪 {name}")
        info = R.get_slm_info(name)
        if info is None:
            st.info("未在信息库中找到该试剂（仍会保存你的选择）")
            return
        img = R.get_slm_photo(name)
        if img:
            st.image(str(img), width="stretch")
        row = info.iloc[0]
        kv = {c: row[c] for c in S.get_display_columns("slm")
              if c in info.columns and row[c] not in (None, "", "None")}
        if kv:
            st.markdown("**关键性质**")
            st.dataframe(pd.DataFrame(list(kv.items()), columns=["属性", "值"]),
                         width="stretch", hide_index=True)
        with st.expander("查看全部属性"):
            st.dataframe(info, width="stretch", hide_index=True)


def page_experiment_list():
    st.title("📋 实验详情")
    conn = get_db()
    exps = M.list_experiments(conn)
    if len(exps) == 0:
        st.info("暂无实验记录。请到「新增实验」录入。")
        return

    st.subheader("选择实验查看详情")
    sel_folder = st.selectbox("选择实验", exps["folder"].tolist(), key="el_sel")
    st.session_state["_cur_detail_folder"] = sel_folder

    # —— 侧栏：分析物 / SLM 详情查询（三页共用，选择跨页保留）——
    def _ph_lookup_det(a):
        try:
            _gcond_d = M.get_group_conditions(conn, int(exps[exps["folder"] == sel_folder]["id"].iloc[0]))
            _phs = pd.to_numeric(_gcond_d["pH_donor"], errors="coerce").dropna()
            return float(_phs.iloc[0]) if len(_phs) else None
        except Exception:
            return None
    row = exps[exps["folder"] == sel_folder].iloc[0]
    eid = int(row["id"])
    raw_labels = M.get_group_labels(conn, eid)
    group_labels = dedup_labels(raw_labels)  # 去重确保列名唯一
    render_detail_sidebar(ph_lookup=_ph_lookup_det)

    # 组级条件数据（供各 tab 使用）
    gcond = M.get_group_conditions(conn, eid)
    for col in ["slm_additives", "donor_additives", "acceptor_additives", "note_tags"]:
        if col in gcond.columns:
            gcond[col] = gcond[col].apply(lambda x: ", ".join(json.loads(x)) if x and x != "[]" else "")
    gcond_display = gcond.copy()
    gcond_display["grp"] = gcond_display["grp"].map(lambda x: "基准" if x == 0 else f"M{x}")
    if "group_label" in gcond_display.columns:
        for i, r in gcond_display.iterrows():
            if r["grp"] != "基准":
                g = int(r["grp"].replace("M", ""))
                if r["group_label"] in (None, "", "未命名SLM") or "未命名" in str(r["group_label"]):
                    gcond_display.at[i, "group_label"] = group_labels.get(g, f"M{g}")
    # 异常组着色（有异常标签的行标红）
    def _has_tags(r):
        return bool(str(r.get("note_tags", "") or "").strip())
    if "note_tags" in gcond_display.columns:
        def _row_style(r):
            return ["background-color: #ffe0e0"] * len(r) if _has_tags(r) else [""] * len(r)
        gcond_display = gcond_display.rename(columns={"note_tags": "异常标签", "well_position": "孔位",
                                                      "group_note": "备注(手动)"})
        _styled_gcond = gcond_display.style.apply(_row_style, axis=1)
    else:
        gcond_display = gcond_display.rename(columns={"well_position": "孔位", "group_note": "备注(手动)"})
        _styled_gcond = gcond_display.style

    mdf = pd.read_sql(f"SELECT * FROM metrics WHERE experiment_id={eid}", conn)
    qc_row = pd.read_sql("SELECT * FROM qc_summary WHERE experiment_id=?", conn, params=(eid,))
    # {组号: 异常标签列表}（供热图组名⚠标记与异常统计）
    grp_tags_map = {}
    for _, gr in gcond.iterrows():
        if gr["grp"] != 0 and str(gr.get("note_tags", "") or "").strip():
            grp_tags_map[int(gr["grp"])] = [t for t in str(gr["note_tags"]).split(", ") if t]

    # 页内导航 tabs
    detail_tabs = st.tabs(["📌 萃取条件", "⚡ 电流曲线", "📊 回收率/转移率热图", "🧊 三相分布与基质效应", "🔧 剔除异常数据", "✏️ 修改/删除"])

    # ---- Tab1: 萃取条件 ----
    with detail_tabs[0]:
        st.markdown("**各组萃取条件**")
        st.dataframe(_styled_gcond, width="stretch", hide_index=True)
        st.caption("淡红色行 = 该组存在异常标签（渗漏/烧焦/破裂等）。")

        # —— 本实验的观察记录（在「📝观察记录」页添加/修改，此处同步展示）——
        _obs_df = M.list_observations(scope="experiment")
        _obs_df = _obs_df[_obs_df["folder"] == sel_folder] if len(_obs_df) else _obs_df
        st.markdown("---")
        st.markdown("**📝 本实验的观察记录**（添加/修改请到「📝 观察记录」页或侧边栏）")
        if len(_obs_df) == 0:
            st.caption("（暂无本实验的观察记录）")
        else:
            for _, _r in _obs_df.iterrows():
                with st.expander(f"{_r['created_at']} | {str(_r['text'])[:40]}"):
                    st.write(_r["text"])
                    _p = M.get_observation_image_path(str(_r["image"]))
                    if _p:
                        st.image(str(_p), width="stretch")

    # ---- Tab2: 电流曲线 ----
    with detail_tabs[1]:
        st.markdown("**电流-时间曲线**")
        exp_type_cur = row.get("exp_type", "")
        # 优先使用用户在新增实验时选择的 sheet 映射
        mapped = None
        try:
            mapped = json.loads(row["current_sheets"]) if row.get("current_sheets") else None
        except Exception:
            mapped = None
        cur_data = {}
        if mapped:
            for m in mapped:
                d = CUR.load_current_pair(m.get("file", ""), m.get("sheet", ""))
                if d:
                    cur_data[f"{m['file']} :: {m['sheet']}"] = d
            if cur_data:
                st.caption("📌 已使用录入实验时选择的电流 sheet 映射")
        if not cur_data:
            # 回退：按命名规则自动匹配
            extract_times = []
            if "extract_time" in gcond.columns:
                extract_times = [t for t in gcond[gcond["grp"] > 0]["extract_time"].dropna().unique().tolist() if t]
            cur_data = CUR.find_current_for_experiment(sel_folder, exp_type_cur, extract_times)
            if cur_data:
                st.caption("📌 按 sheet 命名规则自动匹配（如需固定映射，可在新增实验时选择电流 sheet）")
        if cur_data:
            for sn, data in cur_data.items():
                st.caption(f"📌 数据来源：`{data['source_file']}` > sheet `{sn}`（{data['n_points']} 点，单位 {data['unit']}）")
                st.plotly_chart(CUR.plot_current(data, title=f"{sn} 电流-时间曲线"), width="stretch",
                                key=f"cur_fig_{eid}_{sn}")
        else:
            st.info(f"未找到电流记录。请把电流 xlsx 放入 `data/current/`，"
                    f"或重新录入实验时在第3.5步选择电流 sheet。")

    # ---- Tab3: 回收率/转移率热图 ----
    with detail_tabs[2]:
        if len(mdf):
            # 标准品信息（RSD>15%标红）
            st.markdown("**标准品 / 空白情况**（RSD>15% 标红）")
            if len(qc_row):
                qc_show = qc_row[["analyte", "std_acceptor_mean", "std_acceptor_rsd",
                                  "blk_acceptor_mean", "std_rt_mean", "std_rt_rsd"]].copy()
                qc_show.columns = ["物质", "受体标准品平均值", "受体标准品RSD%",
                                   "受体空白平均值", "标准品保留时间平均值", "保留时间RSD%"]
                rsd_cols = ["受体标准品RSD%", "保留时间RSD%"]
                styled_qc = qc_show.style.map(
                    lambda v: "background-color: #ffc7ce; color: #9C0006"
                    if (isinstance(v, (int, float)) and not pd.isna(v) and v > 15.0) else "", subset=rsd_cols)
                st.dataframe(styled_qc.format("{:.2f}", na_rep="-",
                    subset=[c for c in qc_show.columns if c != "物质"]),
                    width="stretch", hide_index=True)
            st.markdown("**回收率 / 转移率 / 自定义指标 热图**")
            render_heatmaps(mdf, group_labels, qc_row,
                            va=float(row["Va"]), vd=float(row["Vd"]), gcond=gcond,
                            grp_tags_map=grp_tags_map,
                            n_replicates=(int(row["n_replicates"]) if pd.notna(row["n_replicates"]) else 3),
                            excluded=M.get_excluded_samples(conn, eid),
                            folder=sel_folder, eid=eid)
        else:
            st.info("暂无计算结果。")

    # ---- Tab4: 三相分布（供体相/膜/受体相 百分比堆积）----
    with detail_tabs[3]:
        if not bool(row["use_sample"]):
            st.info("本实验未启用供体相（供体相非缓冲液），三相分布不适用。")
        else:
            st.markdown("**三相分布**（各药物在 受体相 / 膜 / 供体相 的百分比堆积）")
            st.caption("供体相%=(供体峰面积−空白)×Vd；受体相%=(受体峰面积−空白)×Va；"
                       "膜相=初始总量−受体相量×校正系数（<0记0）。误差棒=相应相峰面积RSD。")
            try:
                res_ph = _load_calc_res(conn, eid, sel_folder)
                _gqc_ph, _dqc_ph, _ = _build_group_qc_from_db(conn, eid)
                dist = C.compute_phase_distribution(res_ph.df, res_ph.analytes, res_ph.groups,
                                                    float(row["Va"]), float(row["Vd"]),
                                                    excluded_samples=M.get_excluded_samples(conn, eid),
                                                    group_qc=_gqc_ph, default_qc=_dqc_ph)
                if dist:
                    dist_df = pd.DataFrame(dist)
                    # 筛选实验组（可多选：每组一张图并排对比查看）
                    grp_opts = sorted(dist_df["group"].unique())
                    sel_grps_ph = st.multiselect(
                        "选择实验组（可多选，每组一张图对比查看）",
                        [int(g) for g in grp_opts],
                        default=[int(grp_opts[0])],
                        key=f"ph_dist_grp_{eid}",
                        format_func=lambda g: group_labels.get(int(g), f"M{g}"))
                    if not sel_grps_ph:
                        sel_grps_ph = [int(grp_opts[0])]
                    dd = dist_df[dist_df["group"].isin(sel_grps_ph)].copy()
                    # 横坐标 = 药物；每个选中的实验组单独一张图（上下排列方便对比）
                    dd = dd.sort_values(["group", "analyte"])
                    # 配色：受体相紫色(上) / 膜白色(中) / 供体相蓝色(下)
                    for grp in sorted(sel_grps_ph):
                        sub = dd[dd["group"] == grp]
                        gl = group_labels.get(int(grp), f"M{grp}")
                        fig_ph = go.Figure()
                        # 受体相（最上面，紫）
                        fig_ph.add_trace(go.Bar(
                            x=sub["analyte"], y=sub["acceptor_pct"], name="受体相",
                            marker_color="#AB63FA",
                            error_y=dict(type="data",
                                         array=[p * r / 100 if pd.notna(r) else 0
                                                for p, r in zip(sub["acceptor_pct"], sub["acceptor_rsd"])],
                                         visible=True),
                            hovertemplate="%{x}<br>受体相=%{y:.1f}%<extra></extra>"))
                        # 膜（中间，白+灰边）
                        fig_ph.add_trace(go.Bar(
                            x=sub["analyte"], y=sub["membrane_pct"], name="膜",
                            marker_color="white", marker_line=dict(color="#B0B0B0", width=1),
                            hovertemplate="%{x}<br>膜=%{y:.1f}%<extra></extra>"))
                        # 供体相（最下面，蓝）
                        fig_ph.add_trace(go.Bar(
                            x=sub["analyte"], y=sub["donor_pct"], name="供体相",
                            marker_color="#4E79A7",
                            error_y=dict(type="data",
                                         array=[p * r / 100 if pd.notna(r) else 0
                                                for p, r in zip(sub["donor_pct"], sub["donor_rsd"])],
                                         visible=True),
                            hovertemplate="%{x}<br>供体相=%{y:.1f}%<extra></extra>"))
                        fig_ph.update_layout(
                            barmode="stack", title=f"三相分布 — {gl}（受体相/膜/供体相）",
                            xaxis_title="药物", yaxis_title="百分比 (%)",
                            yaxis_range=[0, 100], template="simple_white",
                            height=520, hovermode="x unified",
                            legend=dict(orientation="v", x=1.02, y=1))
                        st.plotly_chart(fig_ph, width="stretch",
                                        key=f"ph_dist_fig_{eid}_{grp}")
                    # 明细表
                    with st.expander("查看三相分布明细数据"):
                        show_d = dd.copy()
                        show_d["group"] = [group_labels.get(int(g), f"M{g}") for g in show_d["group"]]
                        show_d = show_d.rename(columns={"group": "组别", "analyte": "物质",
                                                        "donor_pct": "供体相%", "membrane_pct": "膜%",
                                                        "acceptor_pct": "受体相%",
                                                        "donor_rsd": "供体相RSD%", "acceptor_rsd": "受体相RSD%"})
                        st.dataframe(show_d[["组别", "物质", "供体相%", "膜%", "受体相%",
                                             "供体相RSD%", "受体相RSD%"]],
                                     width="stretch", hide_index=True)

                    # —— 标准品一致性检查（受体相 vs 供体相标准品，50%参考线）——
                    st.markdown("---")
                    st.markdown("**标准品一致性检查**（受体相标准品 vs 供体相标准品峰面积百分比；"
                                "50% 参考线——偏离大 = 两相标准品不一致 / 存在基质效应）")
                    if len(qc_row):
                        qc_std = qc_row[["analyte", "std_acceptor_mean", "std_donor_mean"]].dropna(
                            subset=["std_acceptor_mean", "std_donor_mean"]).copy()
                        qc_std = qc_std[(qc_std["std_acceptor_mean"] > 0) & (qc_std["std_donor_mean"] > 0)]
                        if len(qc_std):
                            qc_std["受体相标准品%"] = qc_std["std_acceptor_mean"] / (
                                qc_std["std_acceptor_mean"] + qc_std["std_donor_mean"]) * 100
                            qc_std["供体相标准品%"] = 100 - qc_std["受体相标准品%"]
                            # 堆叠百分比柱：受体相(黄)在下、供体相(绿)在上，50%线为分界
                            # （供受体同名标准品时各占50%，分界线与50%线重合属正常）
                            fig_std = go.Figure()
                            fig_std.add_trace(go.Bar(
                                x=qc_std["analyte"], y=qc_std["受体相标准品%"], name="受体相标准品",
                                marker_color="#FFC000",
                                hovertemplate="%{x}<br>受体相=%{y:.1f}%<extra></extra>"))
                            fig_std.add_trace(go.Bar(
                                x=qc_std["analyte"], y=qc_std["供体相标准品%"], name="供体相标准品",
                                marker_color="#70AD47",
                                hovertemplate="%{x}<br>供体相=%{y:.1f}%<extra></extra>"))
                            # 50% 参考线：用 trace 画线（数据坐标系，保证可见+图例）
                            fig_std.add_trace(go.Scatter(
                                x=[qc_std["analyte"].iloc[0], qc_std["analyte"].iloc[-1]],
                                y=[50, 50], mode="lines",
                                line=dict(color="red", width=2, dash="dash"),
                                name="50% 参考线（两相一致）"))
                            # 偏离>20%的物质在x轴标注提醒
                            dev_std = (qc_std["受体相标准品%"] - 50).abs()
                            n_dev = int((dev_std > 20).sum())
                            fig_std.update_layout(
                                barmode="stack", title="受体相/供体相标准品峰面积百分比（每物质）",
                                xaxis_title="物质", yaxis_title="标准品均值比 (%)",
                                yaxis_range=[0, 100], template="simple_white",
                                height=480, legend=dict(orientation="h", y=1.12),
                                hovermode="x unified", margin=dict(t=100))
                            st.plotly_chart(fig_std, width="stretch",
                                            key=f"std_check_fig_{eid}")
                            worst_i = dev_std.idxmax()
                            st.caption(f"红虚线=50%（两相标准品完全一致时每物质柱应在50%附近）。"
                                       f"偏离>20%的物质共 {n_dev} 个"
                                       + (f"，最大偏离：{qc_std.loc[worst_i, 'analyte']}"
                                          f"（受体占 {qc_std.loc[worst_i, '受体相标准品%']:.1f}%）——提示可能存在基质效应。"
                                          if n_dev > 0 else "，两相标准品一致性良好。"))
                            # 详情表格
                            with st.expander("查看标准品一致性详情数据"):
                                show_std = qc_std.copy()
                                show_std = show_std.rename(columns={
                                    "analyte": "物质",
                                    "std_acceptor_mean": "受体相标准品峰面积均值",
                                    "std_donor_mean": "供体相标准品峰面积均值"})
                                show_std["偏离50%(百分点)"] = (show_std["受体相标准品%"] - 50).abs().round(1)
                                st.dataframe(show_std[["物质", "受体相标准品峰面积均值", "供体相标准品峰面积均值",
                                                       "受体相标准品%", "供体相标准品%", "偏离50%(百分点)"]],
                                             width="stretch", hide_index=True)
                        else:
                            st.caption("（标准品均值数据不足，无法画一致性图）")
                    else:
                        st.caption("（无标准品数据）")
                else:
                    st.info("无法计算三相分布（标准品数据不足）。")
            except Exception as e:
                st.warning(f"三相分布计算失败：{e}")

            # —— 标准品对比（按名称取同名针平均值，可剔针；表现同基质效应图）——
            st.markdown("---")
            st.markdown("**标准品对比**（选择两个标准品名称，各取**同名进样针的平均值**逐物质对比响应；"
                        "可勾选剔除异常针，图表实时重算；50% 参考线——偏离大 = 两标准品响应不一致）")
            try:
                cmp_long, cmp_cands = _std_compare_data(sel_folder, _folder_sig(sel_folder))
            except Exception as e:
                cmp_long, cmp_cands = pd.DataFrame(), []
                st.caption(f"（标准品对比数据读取失败：{e}）")
            if not cmp_cands:
                st.caption("（数据中无样品名之外的进样（需 ALLDATA 导出），无法对比）")
            else:
                cmp_long["area"] = pd.to_numeric(cmp_long["area"], errors="coerce")
                _name_opts = sorted({t[0] for t in cmp_cands},
                                    key=lambda n: (0 if "std" in n.lower() else 1, n))  # 含 STD 排前
                _needles_by_name = {}
                for _n, _i in cmp_cands:
                    _needles_by_name.setdefault(_n, []).append(_i)

                def _fmt_name(n):
                    return f"⭐ {n}" if "std" in n.lower() else n

                # 默认：受体相标准品 vs 供体相标准品（同名时 B 自动取下一个名称）
                _qc_used_cmp = getattr(res_ph, "qc_names", {}) if "res_ph" in dir() else {}
                _def_na = _qc_used_cmp.get("std_acceptor")
                _def_nb = _qc_used_cmp.get("std_donor")
                _ia = _name_opts.index(_def_na) if _def_na in _name_opts else 0
                _ib = _name_opts.index(_def_nb) if _def_nb in _name_opts else min(1, len(_name_opts) - 1)
                if _ia == _ib:
                    # 两相标准品同名时 B 优先跳到另一个含 STD 的名称，无则下一个名称
                    _alt = next((k for k, n in enumerate(_name_opts)
                                 if k != _ia and "std" in n.lower()), None)
                    _ib = _alt if _alt is not None else (_ia + 1) % len(_name_opts)
                _sc1, _sc2 = st.columns(2)
                _sel_na = _sc1.selectbox("对比标准品 A（同名针平均）", _name_opts, index=_ia,
                                         key=f"stdcmp_a_{eid}", format_func=_fmt_name)
                _sel_nb = _sc2.selectbox("对比标准品 B（同名针平均）", _name_opts, index=_ib,
                                         key=f"stdcmp_b_{eid}", format_func=_fmt_name)

                def _agg_side(side, name):
                    """该名称下：剔针复选框（实时生效）→ 每物质跨针 均值/RSD%。"""
                    needles = sorted(_needles_by_name.get(name, []))
                    excl = set()
                    if needles:
                        _hc = st.columns(min(len(needles), 8) + 1)
                        _hc[0].caption("剔除该针：")
                        for k, nd in enumerate(needles):
                            if _hc[1 + k % 8].checkbox(f"#{nd}", key=f"stdcmp_ex_{eid}_{side}_{name}_{nd}"):
                                excl.add(nd)
                    kept = [n for n in needles if n not in excl]
                    sub = cmp_long[(cmp_long["sample"] == name) & (~cmp_long["injection"].isin(excl))]
                    g = sub.groupby("analyte")["area"]
                    mean_s = g.mean()
                    std_s = g.std(ddof=1)
                    rsd_s = (std_s / mean_s * 100).where(mean_s != 0)
                    return mean_s, rsd_s, kept, excl

                if _sel_na == _sel_nb:
                    st.info("两个名称相同——请选择两个不同的标准品进行对比。")
                else:
                    _ma, _ra, _kept_a, _excl_a = _agg_side("a", _sel_na)
                    _mb, _rb, _kept_b, _excl_b = _agg_side("b", _sel_nb)
                    st.caption("说明：此处的剔针只影响本对比视图；若需在回收率等计算中剔除标准品针，"
                               "请到「新增实验第6步」或「修改/删除」页操作（那里的剔除会保存并参与计算）。")
                    if not _kept_a or not _kept_b:
                        st.warning("有一侧的进样针被全部剔除——请至少保留一根针。")
                    else:
                        def _f(v):
                            try:
                                if v is None or pd.isna(v):
                                    return 0.0
                            except (TypeError, ValueError):
                                return 0.0
                            return float(v)

                        _union = sorted(set(_ma.index) | set(_mb.index))

                        # —— 实时明细表（勾选剔除即时更新；灰=该针已剔除留痕，红=RSD>15%）——
                        _needles_a = sorted(_needles_by_name.get(_sel_na, []))
                        _needles_b = sorted(_needles_by_name.get(_sel_nb, []))
                        _sub_a = cmp_long[cmp_long["sample"] == _sel_na]
                        _sub_b = cmp_long[cmp_long["sample"] == _sel_nb]
                        _area_a = {(int(_nd), _an): _v for _nd, _an, _v in
                                   zip(_sub_a["injection"], _sub_a["analyte"], _sub_a["area"])}
                        _area_b = {(int(_nd), _an): _v for _nd, _an, _v in
                                   zip(_sub_b["injection"], _sub_b["analyte"], _sub_b["area"])}

                        def _num(v):
                            try:
                                if v is None or pd.isna(v):
                                    return None
                                return float(v)
                            except (TypeError, ValueError):
                                return None

                        _live_rows = []
                        for _an in _union:
                            # 注意：变量名避开 row（外层作用域的实验记录同名，覆盖会毁掉后续 tab）
                            _live_row = {"物质": _an}
                            for _nd in _needles_a:
                                _live_row[f"A#{_nd}"] = _num(_area_a.get((_nd, _an)))
                            _live_row["A均值"] = _num(_ma.get(_an))
                            _live_row["A RSD%"] = _num(_ra.get(_an))
                            for _nd in _needles_b:
                                _live_row[f"B#{_nd}"] = _num(_area_b.get((_nd, _an)))
                            _live_row["B均值"] = _num(_mb.get(_an))
                            _live_row["B RSD%"] = _num(_rb.get(_an))
                            _live_rows.append(_live_row)
                        _live = pd.DataFrame(_live_rows).set_index("物质")
                        _excl_cols = {f"A#{_nd}" for _nd in _excl_a} | {f"B#{_nd}" for _nd in _excl_b}

                        def _live_style(s):
                            styles = []
                            for c in s.index:
                                if c in _excl_cols:
                                    styles.append("background-color: #d9d9d9; color: #808080")
                                elif "RSD%" in c and pd.notna(s[c]) and float(s[c]) > 15.0:
                                    styles.append("background-color: #ffc7ce; color: #9C0006")
                                else:
                                    styles.append("")
                            return styles

                        st.dataframe(_live.style.apply(_live_style, axis=1)
                                     .format({c: "{:.2e}" for c in _live.columns if "RSD" not in c}
                                             | {c: "{:.2f}" for c in _live.columns if "RSD" in c},
                                             na_rep="-"),
                                     width="stretch")
                        st.caption("明细表实时更新：灰=该针已剔除（显示原值留痕）；红=RSD>15%；均值/RSD 只计未剔除的针。")

                        _rows_cmp = []
                        for _an in _union:
                            _x = _f(_ma.get(_an))
                            _y = _f(_mb.get(_an))
                            if _x + _y <= 0:
                                continue
                            _rows_cmp.append({
                                "analyte": _an, "面积A": _x, "面积B": _y,
                                "A%": _x / (_x + _y) * 100.0, "B%": _y / (_x + _y) * 100.0,
                                "RSD_A": _f(_ra.get(_an)), "RSD_B": _f(_rb.get(_an)),
                            })
                        if not _rows_cmp:
                            st.caption("（两个标准品无共同响应数据）")
                        else:
                            cmp_df = pd.DataFrame(_rows_cmp)
                            _lab_a = f"{_fmt_name(_sel_na)}（{len(_kept_a)}针平均）"
                            _lab_b = f"{_fmt_name(_sel_nb)}（{len(_kept_b)}针平均）"
                            fig_cmp = go.Figure()
                            fig_cmp.add_trace(go.Bar(
                                x=cmp_df["analyte"], y=cmp_df["A%"], name=_lab_a,
                                marker_color="#FFC000", customdata=cmp_df["RSD_A"].to_numpy(),
                                hovertemplate="%{x}<br>" + _lab_a + "=%{y:.1f}%"
                                              "<br>RSD=%{customdata:.1f}%<extra></extra>"))
                            fig_cmp.add_trace(go.Bar(
                                x=cmp_df["analyte"], y=cmp_df["B%"], name=_lab_b,
                                marker_color="#70AD47", customdata=cmp_df["RSD_B"].to_numpy(),
                                hovertemplate="%{x}<br>" + _lab_b + "=%{y:.1f}%"
                                              "<br>RSD=%{customdata:.1f}%<extra></extra>"))
                            # 50% 参考线：trace 画线（数据坐标系）
                            fig_cmp.add_trace(go.Scatter(
                                x=[cmp_df["analyte"].iloc[0], cmp_df["analyte"].iloc[-1]],
                                y=[50, 50], mode="lines",
                                line=dict(color="red", width=2, dash="dash"),
                                name="50% 参考线（两标准品一致）"))
                            _dev_cmp = (cmp_df["A%"] - 50).abs()
                            _n_dev_cmp = int((_dev_cmp > 20).sum())
                            _worst_cmp = _dev_cmp.idxmax()
                            fig_cmp.update_layout(
                                barmode="stack",
                                title=f"标准品对比：{_lab_a} vs {_lab_b}（每物质响应占比）",
                                xaxis_title="物质", yaxis_title="占比 (%)",
                                yaxis_range=[0, 100], template="simple_white",
                                height=480, legend=dict(orientation="h", y=1.12),
                                hovermode="x unified", margin=dict(t=100))
                            st.plotly_chart(fig_cmp, width="stretch",
                                            key=f"stdcmp_fig_{eid}")
                            st.caption(f"红虚线=50%（两标准品响应完全一致时每物质柱应在50%附近）。"
                                       f"偏离>20%的物质共 {_n_dev_cmp} 个"
                                       + (f"，最大偏离：{cmp_df.loc[_worst_cmp, 'analyte']}"
                                          f"（A 占 {cmp_df.loc[_worst_cmp, 'A%']:.1f}%）。"
                                          if _n_dev_cmp > 0 else "，两标准品响应一致性良好。"))
                            with st.expander("查看标准品对比详情数据"):
                                show_cmp = cmp_df.rename(columns={
                                    "analyte": "物质",
                                    "面积A": f"{_lab_a} 峰面积", "面积B": f"{_lab_b} 峰面积",
                                    "A%": f"{_lab_a} %", "B%": f"{_lab_b} %",
                                    "RSD_A": f"{_lab_a} RSD%", "RSD_B": f"{_lab_b} RSD%"})
                                show_cmp["偏离50%(百分点)"] = (show_cmp[f"{_lab_a} %"] - 50).abs().round(1)
                                st.dataframe(
                                    show_cmp[["物质", f"{_lab_a} 峰面积", f"{_lab_a} RSD%",
                                              f"{_lab_b} 峰面积", f"{_lab_b} RSD%",
                                              f"{_lab_a} %", f"{_lab_b} %", "偏离50%(百分点)"]],
                                    width="stretch", hide_index=True)

    # ---- Tab5: 剔除异常数据 ----
    with detail_tabs[4]:
        if len(mdf):
            st.markdown("勾选要剔除的样本，重算后热图实时更新。")
            _exclusion_ui_fragment(eid, sel_folder, group_labels)
        else:
            st.info("暂无计算结果，无法剔除。")

    # ---- Tab6: 修改/删除 ----
    with detail_tabs[5]:
        st.markdown("**✏️ 修改实验条件**（改后重新生成 DATA.xlsx）")
        _edit_experiment_fragment(eid, row, sel_folder)
        st.markdown("---")
        st.markdown("**⚠️ 删除实验记录**")
        _del_opts = exps["folder"].tolist()
        del_folder = st.selectbox("选择删除", _del_opts,
                                  index=_del_opts.index(sel_folder) if sel_folder in _del_opts else 0,
                                  key="del")
        if st.button("确认删除", type="secondary"):
            did = int(exps[exps["folder"] == del_folder]["id"].iloc[0])
            M.delete_experiment(conn, did)
            # 清该文件夹的会话解析缓存（内含旧实验的标准品剔除状态）：同文件夹复录时防止复用
            st.session_state.pop(f"_excl_res_cache", None)
            st.session_state.pop(f"_excl_res_cache_sig", None)
            _gdo_clear = getattr(globals().get("_global_detail_options"), "clear", None)
            if _gdo_clear:
                try:
                    _gdo_clear()
                except Exception:
                    pass
            st.session_state["_flash_msg"] = f"已删除实验 {del_folder}。"
            st.rerun()


def _fmt_stat_aware(v):
    """Styler 格式器：统计行整数显示为整数（11 而非 11.00），数据行保留两位小数。
    依赖调用前把统计行值转成 int（float 值仍按两位小数显示）。"""
    if pd.isna(v):
        return "-"
    if isinstance(v, int) and not isinstance(v, bool):
        return f"{v:d}"
    return f"{v:.2f}"


def render_heatmaps(mdf: pd.DataFrame, group_labels: dict, qc_df_row: pd.DataFrame,
                    va: float = 50.0, vd: float = 250.0, gcond: pd.DataFrame = None,
                    grp_tags_map: dict = None,
                    n_replicates: int = 3, excluded: set = None,
                    folder: str = "", eid=None):
    """三个热图（各独立色标+统计行）+ 自定义指标热图 + 两个峰面积RSD表。组别名已去重。
    grp_tags_map: {组号: 异常标签列表}——异常组的列名追加 ⚠标记。
    n_replicates/excluded: 回收率热图底部的「有效膜数」= 重复数 − 该组被剔除样本数。
    """
    # 全局配色（所有热图共用；自动记住上次选择，下次打开默认为该配色）
    cmap_options = ["RdYlGn", "viridis", "plasma", "Blues", "RdBu", "YlGnBu", "coolwarm"]
    _cmap_pref_hm = S.get_cmap_pref()
    cmap_use = st.selectbox("配色（所有热图共用，自动记住上次选择）", cmap_options,
                            index=cmap_options.index(_cmap_pref_hm) if _cmap_pref_hm in cmap_options else 0,
                            key="hm_cmap",
                            on_change=lambda: S.save_cmap_pref(st.session_state.get("hm_cmap", "RdYlGn")))

    grp_tags_map = grp_tags_map or {}
    # 原始数据（供峰高/信噪比现算）：只取【当前实验文件夹】的缓存；缺失时按需加载。
    # 旧实现遍历缓存取第一个有 df 的条目（ParseResult 无 folder 属性 → 恒 AttributeError
    # 被吞掉 → 退回 cur_res = 新增实验页最后解析的实验），会拿错实验的数据。
    res_detail_df = None
    if folder:
        _v_hm = st.session_state.get("_excl_res_cache", {}).get(folder)
        if _v_hm is not None and getattr(_v_hm, "df", None) is not None:
            res_detail_df = _v_hm.df
        else:
            try:
                res_detail_df = _load_calc_res(get_db(), eid, folder).df
            except Exception:
                res_detail_df = None

    def _glabel(g):
        base_lbl = group_labels.get(int(g), f"M{g}")
        tags = grp_tags_map.get(int(g))
        if tags:
            return f"{base_lbl} ⚠{'/'.join(tags)}"
        return base_lbl

    metric_cols = [("recovery", "回收率%", "hm_rec", "recovery"),
                   ("dm", "供体相至SLM转移率%", "hm_dm", "transfer"),
                   ("ma", "SLM至受体相转移率%", "hm_ma", "transfer")]

    for col, title, kpre, ref_kind in metric_cols:
        st.markdown(f"#### {title}")
        sub = mdf[["analyte", "grp", col]].copy()
        if sub[col].isna().all():
            st.caption("（无数据，可能未启用供体相）")
            continue
        pivot = sub.pivot_table(index="analyte", columns="grp", values=col, aggfunc="first")
        pivot.columns = [_glabel(c) for c in pivot.columns]
        n_analytes = len(pivot)  # 数据行数（统计行追加前记录，色阶只作用于这些行）
        # —— 独立色标（默认按本指标数据的范围）——
        data_vals = pivot.values.flatten()
        data_vals = data_vals[~pd.isna(data_vals)]
        dmin = float(np.floor(data_vals.min())) if len(data_vals) else 0.0
        dmax = float(np.ceil(data_vals.max())) if len(data_vals) else 100.0
        rc = st.columns([2, 1, 1, 2])
        auto_range = rc[0].checkbox("自动色标范围", value=False, key=f"{kpre}_auto")
        if auto_range:
            vmin_use, vmax_use = None, None
            range_desc = f"自动（{dmin:g} ~ {dmax:g}）"
        else:
            vmin_use = rc[1].number_input("颜色最小值", value=dmin, key=f"{kpre}_vmin")
            vmax_use = rc[2].number_input("颜色最大值", value=dmax, key=f"{kpre}_vmax")
            range_desc = f"{vmin_use:g} ~ {vmax_use:g}"
        # —— 参考线（默认值按指标分开存储；改动后可保存为该指标的默认）——
        dflt_r1, dflt_r2 = S.get_ref_lines(ref_kind)
        with st.expander("📐 统计参考线设置（决定底部统计行的分档）"):
            # 转移率默认未设置参考线 → 启用与否由用户勾选；回收率默认启用（有默认值时）
            use_ref = st.checkbox("启用统计参考线",
                                  value=(dflt_r1 is not None), key=f"{kpre}_ref_on")
            ref1 = ref2 = None
            if use_ref:
                rr1, rr2 = st.columns(2)
                ref1 = rr1.number_input("参考线1", value=float(dflt_r1 if dflt_r1 is not None else 20.0),
                                        key=f"{kpre}_ref1")
                use_ref2 = rr2.checkbox("启用第2条参考线", value=dflt_r2 is not None,
                                        key=f"{kpre}_ref2_on")
                if use_ref2:
                    ref2 = rr2.number_input("参考线2", value=float(dflt_r2 if dflt_r2 is not None else 40.0),
                                            key=f"{kpre}_ref2")
                if st.button("💾 保存为该指标默认参考线", key=f"{kpre}_ref_save",
                             help=f"保存后「⚙️ 设置」页与各处统计的默认值同步更新（{ref_kind}）"):
                    S.save_ref_lines(ref1 if ref1 is not None else dflt_r1,
                                     ref2 if ref2 is not None else dflt_r2,
                                     kind=ref_kind)
                    st.success("已保存为该指标默认参考线。")
        styled = pivot.style.format("{:.2f}", na_rep="-")
        # —— 统计行：按列（每个实验组）统计各档个数 ——
        name_short = title.split("%")[0]
        if use_ref and ref1 is not None:
            # 分档：线2存在时为 [<线1, 线1~线2, ≥线2]；只有线1时为 [<线1, ≥线1]
            if ref2 is not None and ref2 > ref1:
                lo, hi = ref1, ref2
                row_lo = [int(((pivot[c] < lo)).sum()) for c in pivot.columns]
                row_mid = [int((((pivot[c] >= lo) & (pivot[c] < hi))).sum()) for c in pivot.columns]
                row_hi = [int(((pivot[c] >= hi)).sum()) for c in pivot.columns]
                pivot.loc[f"📊 <{lo:g} 个数"] = row_lo
                pivot.loc[f"📊 {lo:g}~{hi:g} 个数"] = row_mid
                pivot.loc[f"📊 ≥{hi:g} 个数"] = row_hi
                stat_desc = f"按实验组统计：<{lo:g} / {lo:g}~{hi:g} / ≥{hi:g} 的个数"
            else:
                row_lo = [int(((pivot[c] < ref1)).sum()) for c in pivot.columns]
                row_hi = [int(((pivot[c] >= ref1)).sum()) for c in pivot.columns]
                pivot.loc[f"📊 <{ref1:g} 个数"] = row_lo
                pivot.loc[f"📊 ≥{ref1:g} 个数"] = row_hi
                stat_desc = f"按实验组统计：<{ref1:g} / ≥{ref1:g} 的个数"
            if col == "recovery":
                # 回收率热图始终显示有效膜数行
                excluded = excluded or set()
                row_valid = []
                for c_g in pivot.columns:
                    g_num = next((gn for gn, lbl in group_labels.items()
                                  if lbl == c_g or c_g.startswith(lbl + " ⚠")), None)
                    n_excl = sum(1 for s_ in excluded
                                 if (P.parse_sample_name(str(s_)) or (None,))[0] == g_num) \
                        if g_num is not None else 0
                    row_valid.append(max(0, n_replicates - n_excl))
                pivot.loc["📊 有效膜数"] = row_valid
                stat_desc += "；底部末行为各实验组有效膜数（重复数−被剔除数）"
        else:
            if col == "recovery":
                # 有效膜数 = 重复数 − 该组被剔除样本数（仅回收率显示）
                excluded = excluded or set()
                row_valid = []
                for c_g in pivot.columns:
                    # 列名可能带 ⚠ 后缀，从 group_labels 反查组号
                    g_num = next((gn for gn, lbl in group_labels.items()
                                  if lbl == c_g or c_g.startswith(lbl + " ⚠")), None)
                    n_excl = sum(1 for s_ in excluded
                                 if (P.parse_sample_name(str(s_)) or (None,))[0] == g_num) \
                        if g_num is not None else 0
                    row_valid.append(max(0, n_replicates - n_excl))
                pivot.loc["📊 有效膜数"] = row_valid
                stat_desc = "未启用参考线。底部行为各实验组有效膜数（重复数−被剔除数）"
            else:
                stat_desc = "未启用参考线（可在上方参考线设置中启用以显示分档统计）"
        _stat_rows_hm = [r for r in pivot.index if str(r).startswith("📊")]
        if _stat_rows_hm:
            pivot.loc[_stat_rows_hm] = pivot.loc[_stat_rows_hm].fillna(0).astype(int)
        # 统计行整数显示（11 而非 11.00）；gmap 传数值矩阵，兼容 object 混合帧的色阶
        styled = pivot.style.format(_fmt_stat_aware, na_rep="-")
        styled = styled.background_gradient(cmap=cmap_use, axis=None,
                                            vmin=vmin_use, vmax=vmax_use,
                                            gmap=pivot.astype(float),
                                            subset=pd.IndexSlice[pivot.index[:n_analytes], :])
        st.dataframe(styled, width="stretch")
        st.caption(f"颜色范围：{range_desc}（配色 {cmap_use}）。{stat_desc}。统计行白底不参与颜色映射。")

    # —— 三个热图后：供体/受体峰面积RSD两个表（>15%标红）——
    st.markdown("---")
    st.markdown("**峰面积 RSD 表**（>15% 标红，推荐剔除对应偏离样本）")
    for rsd_col, rsd_title in [("A_rsd", "受体峰面积 RSD%"), ("S_rsd", "供体峰面积 RSD%")]:
        st.markdown(f"#### {rsd_title}")
        sub_r = mdf[["analyte", "grp", rsd_col]].copy()
        if sub_r[rsd_col].isna().all():
            st.caption("（无数据）"); continue
        pivot_r = sub_r.pivot_table(index="analyte", columns="grp", values=rsd_col, aggfunc="first")
        pivot_r.columns = [_glabel(c) for c in pivot_r.columns]
        def _hl(v):
            if pd.isna(v): return ""
            try:
                return "background-color: #ffc7ce; color: #9C0006" if float(v) > 15.0 else ""
            except (ValueError, TypeError):
                return ""
        # 统计行：RSD>15% 个数（按列统计）
        count_row = (pivot_r > 15).sum()
        pivot_r.loc["📊 RSD>15% 个数"] = count_row
        pivot_r.loc["📊 RSD>15% 个数"] = pivot_r.loc["📊 RSD>15% 个数"].fillna(0).astype(int)
        styled_r = pivot_r.style.map(_hl).format(_fmt_stat_aware, na_rep="-")
        st.dataframe(styled_r, width="stretch")
        st.caption(f"红色 = {rsd_title} > 15%，推荐考虑剔除该组对应相的偏离样本")

    # —— 峰高热图（处理方式同峰面积：受体相均值+RSD；>15% RSD 标红）——
    st.markdown("---")
    st.markdown("**峰高**（与峰面积同类处理：均值 + RSD；>15% RSD 标红）")
    if res_detail_df is not None:
        def _agg_matrix(value_col: str):
            rows_h = []
            for analyte in sorted(res_detail_df["analyte"].unique()):
                for g in sorted(res_detail_df["group"].dropna().unique()):
                    g = int(g)
                    sub = res_detail_df[(res_detail_df["category"] == "sample") &
                                        (res_detail_df["analyte"] == analyte) &
                                        (res_detail_df["group"] == g) &
                                        (res_detail_df["phase"] == "A")]
                    vals = [v for v in sub[value_col].tolist()
                            if v and not (isinstance(v, float) and v != v)]
                    if vals:
                        import numpy as _np
                        m = float(_np.mean(vals))
                        rsd = (float(_np.std(vals, ddof=1)) / m * 100) if (len(vals) >= 2 and m) else None
                        rows_h.append({"analyte": analyte, "grp": g, "mean": m, "rsd": rsd})
            return pd.DataFrame(rows_h)

        hm_h = _agg_matrix("height")
        # 手动颜色映射（与汇总表同方案；background_gradient 在暗色主题下会渲染成黑块）
        import matplotlib as _mpl
        _cmap_h = _mpl.colormaps["RdYlGn"]
        def _colored_table(pv: pd.DataFrame, cmap_obj, na_text="-"):
            vals = pv.values.flatten()
            vals = vals[~pd.isna(vals)]
            lo = float(np.min(vals)) if len(vals) else 0.0
            hi = float(np.max(vals)) if len(vals) else 1.0
            if lo == hi:
                hi = lo + 1.0
            norm = _mpl.colors.Normalize(vmin=lo, vmax=hi)
            disp = pv.copy().astype(object)
            for r in pv.index:
                for c in pv.columns:
                    v = pv.loc[r, c]
                    if pd.isna(v):
                        disp.loc[r, c] = na_text
                    else:
                        disp.loc[r, c] = f"{float(v):.3e}"
            def _style_row(s):
                out_ = []
                for c in s.index:
                    v = pv.loc[s.name, c]
                    if pd.isna(v):
                        out_.append("")
                    else:
                        out_.append(f"background-color: {_mpl.colors.to_hex(cmap_obj(norm(float(v))))}")
                return out_
            return disp.style.apply(_style_row, axis=1), lo, hi

        if len(hm_h):
            st.markdown("#### 峰高平均值（受体相，a.u.；颜色=数值映射）")
            pv = hm_h.pivot_table(index="analyte", columns="grp", values="mean", aggfunc="first")
            pv.columns = [_glabel(c) for c in pv.columns]
            styled_h, hlo, hhi = _colored_table(pv, _cmap_h)
            st.dataframe(styled_h, width="stretch")
            st.caption(f"色标：{hlo:.3e} ~ {hhi:.3e}")
            st.markdown("#### 峰高RSD%（>15% 标红）")
            pr = hm_h.pivot_table(index="analyte", columns="grp", values="rsd", aggfunc="first")
            pr.columns = [_glabel(c) for c in pr.columns]
            def _hl_h(v):
                if pd.isna(v): return ""
                try:
                    return "background-color: #ffc7ce; color: #9C0006" if float(v) > 15.0 else ""
                except (ValueError, TypeError):
                    return ""
            cnt_h = (pr > 15).sum()
            pr.loc["📊 RSD>15% 个数"] = cnt_h
            pr.loc["📊 RSD>15% 个数"] = pr.loc["📊 RSD>15% 个数"].fillna(0).astype(int)
            st.dataframe(pr.style.map(_hl_h).format(_fmt_stat_aware, na_rep="-"),
                         width="stretch")
        else:
            st.caption("（无峰高数据）")
        # 信噪比：从 ALLDATA 的 SN 列读取（parser 已提取为 sn 列；旧版 ALLDATA 无该列时提示）
        st.markdown("#### 信噪比（SN，来自 ALLDATA.txt 的 SN 列）")
        if res_detail_df is None or "sn" not in getattr(res_detail_df, "columns", []):
            st.caption("（ALLDATA.txt 中未包含 SN 列——请质谱导出时勾选信噪比，与 ALLDATA 一起放入实验文件夹）")
        else:
            sub_sn = res_detail_df[(res_detail_df["category"] == "sample") &
                                   (res_detail_df["phase"] == "A")][["analyte", "group", "sn"]].copy()
            sub_sn = sub_sn[sub_sn["sn"] > 0]
            sub_sn["grp"] = sub_sn["group"].astype(int)
            if len(sub_sn) == 0:
                st.caption("（SN 列无可解析的受体相样本数据）")
            else:
                agg_sn = sub_sn.groupby(["analyte", "grp"])["sn"].agg(
                    mean="mean",
                    rsd=lambda x: (x.std(ddof=1)/x.mean()*100) if (len(x) >= 2 and x.mean()) else None
                ).reset_index()
                st.markdown("**SN平均值（受体相；颜色=数值映射）**")
                pv_sn = agg_sn.pivot_table(index="analyte", columns="grp", values="mean", aggfunc="first")
                pv_sn.columns = [_glabel(c) for c in pv_sn.columns]
                styled_sn, slo, shi = _colored_table(pv_sn, _cmap_h)
                st.dataframe(styled_sn, width="stretch")
                st.caption(f"色标：{slo:.3e} ~ {shi:.3e}")
                st.markdown("**SN RSD%（>15% 标红）**")
                pr_sn = agg_sn.pivot_table(index="analyte", columns="grp", values="rsd", aggfunc="first")
                pr_sn.columns = [_glabel(c) for c in pr_sn.columns]
                cnt_sn = (pr_sn > 15).sum()
                pr_sn.loc["📊 RSD>15% 个数"] = cnt_sn
                st.dataframe(pr_sn.style.map(_hl_h).format("{:.2f}", na_rep="-"),
                             width="stretch")
    else:
        st.caption("（需要原始数据以计算峰高/信噪比）")

    # —— 自定义指标热图（用户在「🧮 公式说明」中定义）——
    custom_fs = S.get_custom_formulas()
    if custom_fs:
        st.markdown("---")
        st.markdown("**自定义指标**（在「🧮 公式说明 → 自定义公式」中定义）")
        # 合并标准品/空白均值（来自 qc_summary）到每行
        base_cols = ["analyte", "grp", "A_mean", "S_mean", "corr"]
        mdf_c = mdf[base_cols].copy()
        if len(qc_df_row):
            qc_map = qc_df_row[["analyte", "std_acceptor_mean", "std_donor_mean",
                                "blk_acceptor_mean", "blk_donor_mean"]].set_index("analyte")
            for c in ["std_acceptor_mean", "std_donor_mean", "blk_acceptor_mean", "blk_donor_mean"]:
                mdf_c[c] = mdf_c["analyte"].map(qc_map[c])
        else:
            for c in ["std_acceptor_mean", "std_donor_mean", "blk_acceptor_mean", "blk_donor_mean"]:
                mdf_c[c] = None
        # 每组 pH（来自组级条件）
        ph_map = {}
        if gcond is not None and len(gcond):
            for _, gr in gcond.iterrows():
                if gr["grp"] != 0:
                    ph_map[int(gr["grp"])] = (gr.get("pH_donor"), gr.get("pH_acceptor"))
        # 逐行计算
        for cf in custom_fs:
            col_name = cf["name"]
            vals = []
            _ferrs = []  # 求值失败不再静默丢弃：汇总一行原因，用户可排查
            for _, r in mdf_c.iterrows():
                g = int(r["grp"]) if pd.notna(r["grp"]) else 0
                ph_d, ph_a = ph_map.get(g, (None, None))
                variables = {
                    "A_mean": r["A_mean"] or 0.0, "S_mean": r["S_mean"] or 0.0,
                    "STD_acceptor_mean": r["std_acceptor_mean"] or 0.0,
                    "STD_donor_mean": r["std_donor_mean"] or 0.0,
                    "blk_acceptor_mean": r["blk_acceptor_mean"] or 0.0,
                    "blk_donor_mean": r["blk_donor_mean"] or 0.0,
                    "Va": va, "Vd": vd,
                    "pH_donor": float(ph_d) if ph_d else 0.0,
                    "pH_acceptor": float(ph_a) if ph_a else 0.0,
                    "corr": r["corr"] if pd.notna(r["corr"]) else 1.0,
                }
                v, err = C.eval_custom_formula(cf["expr"], variables)
                if err:
                    _ferrs.append(f"{r['analyte']}(M{g})：{err}")
                vals.append(v)
            mdf_c[col_name] = vals
            if _ferrs:
                with st.expander(f"⚠️ {col_name}：{len(_ferrs)} 行计算失败（点击查看原因）"):
                    for _fe in _ferrs[:12]:
                        st.caption(_fe)
            # 画热图
            if mdf_c[col_name].isna().all():
                st.caption(f"{col_name}：无有效计算结果"); continue
            st.markdown(f"#### {col_name}")
            sub_c = mdf_c[["analyte", "grp", col_name]]
            pivot_c = sub_c.pivot_table(index="analyte", columns="grp", values=col_name, aggfunc="first")
            pivot_c.columns = [_glabel(c) for c in pivot_c.columns]
            dv = pivot_c.values.flatten(); dv = dv[~pd.isna(dv)]
            cmin_c = float(np.floor(dv.min())) if len(dv) else 0.0
            cmax_c = float(np.ceil(dv.max())) if len(dv) else 100.0
            styled_c = pivot_c.style.format("{:.2f}", na_rep="-").background_gradient(
                cmap=cmap_use, axis=None, vmin=cmin_c, vmax=cmax_c)
            st.dataframe(styled_c, width="stretch")
            st.caption(f"色标（自动）：{cmin_c:g} ~ {cmax_c:g}")


def _as_name_list(v):
    """添加剂字段容错转列表：list 原样 / JSON 字符串解析 / 其他→单元素或空。
    （快速表格的会话状态里是 list，数据库/自动填入里是 JSON 串，两者都要支持）"""
    if v is None or v == "" or v == "[]":
        return []
    if isinstance(v, (list, tuple)):
        return [str(x) for x in v]
    try:
        return [str(x) for x in json.loads(v)]
    except Exception:
        return [str(v)]


QC_ROLES = [("std_acceptor", "受体相标准品"), ("std_donor", "供体相标准品"),
            ("blk_acceptor", "受体相空白"), ("blk_donor", "供体相空白")]


def _render_qc_combo_editor(all_samples: list, combos: list, key_prefix: str) -> list:
    """编辑多标准品组合（名称 + 四个角色下拉）。返回组合 list[dict]。
    每个组合：{name, std_acceptor, std_donor, blk_acceptor, blk_donor}。"""
    new_combos = []
    for _i, _c in enumerate(combos):
        with st.expander(f"组合 {_i + 1}：{_c.get('name') or '（未命名）'}",
                         expanded=(_i == len(combos) - 1)):
            _nm = st.text_input("组合名称", value=str(_c.get("name") or ""),
                                key=f"{key_prefix}_combo_name_{_i}")
            _roles = {}
            _rc = st.columns(2)
            for _ri, (_role, _rlabel) in enumerate(QC_ROLES):
                _cur = _c.get(_role) or ""
                _idx = all_samples.index(_cur) if _cur in all_samples else 0
                _roles[_role] = _rc[_ri % 2].selectbox(
                    _rlabel, all_samples, index=_idx,
                    key=f"{key_prefix}_combo_{_i}_{_role}")
            new_combos.append({"name": _nm.strip() or f"组合{_i + 1}", **_roles})
    if st.button("➕ 添加一组标准品组合", key=f"{key_prefix}_combo_add"):
        new_combos.append({"name": f"组合{len(combos) + 1}",
                           "std_acceptor": all_samples[0] if all_samples else "",
                           "std_donor": all_samples[0] if all_samples else "",
                           "blk_acceptor": all_samples[0] if all_samples else "",
                           "blk_donor": all_samples[0] if all_samples else ""})
        st.session_state[f"{key_prefix}_qc_combos"] = new_combos
        st.rerun()
    return new_combos


def _combos_union_qc(combos: list) -> dict:
    """组合列表 → 各角色的名字并集 {role: [names]}（供解析器保留全部行）。"""
    union = {role: [] for role, _ in QC_ROLES}
    for _c in combos or []:
        for role, _ in QC_ROLES:
            v = _c.get(role)
            if v and v not in union[role]:
                union[role].append(v)
    return union


def _build_group_qc_from_db(conn, eid: int):
    """从数据库读 (group_qc, default_qc, combos)。
    group_qc: {组号: {role: 名}}；default_qc: 默认组合（第一个/基础质控名）。"""
    row = conn.execute("SELECT qc_names, qc_combos FROM experiments WHERE id=?", (eid,)).fetchone()
    try:
        base_qn = json.loads(row["qc_names"]) if row and row["qc_names"] else {}
    except Exception:
        base_qn = {}
    combos = []
    if row and row["qc_combos"]:
        try:
            combos = json.loads(row["qc_combos"])
        except Exception:
            combos = []
    default_qc = ({k: (v[0] if isinstance(v, list) else v) for k, v in base_qn.items()}
                  if base_qn else
                  ({k: combos[0].get(k) for k, _ in QC_ROLES} if combos else {}))
    if combos:
        default_qc = {k: combos[0].get(k) or default_qc.get(k) for k, _ in QC_ROLES}
    by_name = {c.get("name"): c for c in combos}
    gq = {}
    for r in conn.execute("SELECT grp, qc_combo FROM group_conditions WHERE experiment_id=? AND grp>0", (eid,)):
        cn = r["qc_combo"]
        if cn and cn in by_name:
            gq[int(r["grp"])] = {k: by_name[cn].get(k) for k, _ in QC_ROLES}
    return gq, default_qc, combos


def _sample_example_for(label: str) -> str:
    """按格式选择给出对应的示例样品名（预览框联动）。"""
    if "组-相" in label:
        return "1-A-1"
    if "相组" in label:
        return "A1-1"
    if "自定义" in label:
        return "1-A-1 或 A1-1"
    return "1-A-1"


def _base_ratio_map(entries) -> dict:
    """从基准添加剂条目（可能带 "名称 (5%)" 后缀）提取 {裸名: 比例} 映射，
    供逐组含量输入框继承基准已设比例。"""
    out = {}
    for e in (entries or []):
        c, p = R.split_name_ratio(e)
        if c not in out:
            out[c] = p
    return out


def _expand_std_link(pairs_by_phase: dict, names_by_phase: dict) -> dict:
    """剔除联动（同名标准品才联动）：
    一相勾选了某标准品名 A-STD 的第 i 针 → 另一相【同名】A-STD 的第 i 针一并剔除
    （同一物理进样在两相各有一份行）；不同名的标准品（B-STD 等）不受影响——
    多标准品组合下各溶液独立，剔除按"标准品名+针号"全局生效。
    pairs_by_phase: {相: [(样名, 针号), ...]}；names_by_phase: {相: {样名, ...}}
    """
    out = {k: set(v) for k, v in pairs_by_phase.items()}
    for src, dst in [("std_acceptor", "std_donor"), ("std_donor", "std_acceptor")]:
        for (s, i) in pairs_by_phase.get(src, []):
            if s in names_by_phase.get(dst, set()):
                out.setdefault(dst, set()).add((s, i))
    return out


def _load_calc_res(conn, eid: int, folder: str):
    """重新解析实验并套用已保存的标准品进样针剔除，供所有重算路径统一使用。
    （解析时传入该实验保存的 qc_names，避免自定义标准品/空白名丢失导致质控行识别失败）
    """
    qrow = conn.execute("SELECT qc_names, sample_format, qc_combos FROM experiments WHERE id=?", (eid,)).fetchone()
    try:
        qc_names = json.loads(qrow["qc_names"]) if qrow and qrow["qc_names"] else {}
    except Exception:
        qc_names = {}
    try:
        _combos_loaded = json.loads(qrow["qc_combos"]) if qrow and qrow["qc_combos"] else []
    except Exception:
        _combos_loaded = []
    # 多标准品组合：解析需保留全部组合的标准品/空白行（名字并集）
    if _combos_loaded:
        qc_names = _combos_union_qc(_combos_loaded)
    _qc_json_ne = json.dumps(qc_names, ensure_ascii=False) if qc_names else ""
    _sig_ne = _folder_sig(folder)
    _fmt_stored = (qrow["sample_format"] if qrow else None) or None
    res = _cached_parse(folder, _qc_json_ne, _fmt_stored or "", _sig_ne)
    if not res.groups:
        # 自愈回退：该实验没存格式（旧记录）或存的全局格式与数据命名不符时，
        # 依次尝试各预设格式，找到能识别出样品组的那个并回存到该实验
        for _lbl_f, _rx_f in S.SAMPLE_FORMAT_PRESETS.items():
            if _rx_f == (_fmt_stored or ""):
                continue
            _res_try = _cached_parse(folder, _qc_json_ne, _rx_f, _sig_ne)
            if _res_try.groups:
                res = _res_try
                try:
                    conn.execute("UPDATE experiments SET sample_format=? WHERE id=?", (_rx_f, eid))
                    conn.commit()
                except Exception:
                    pass
                break
    std_excl = M.get_excluded_std(conn, eid)
    drop_pairs = set()
    for pairs in std_excl.values():
        drop_pairs |= pairs
    if drop_pairs:
        mask = res.df.apply(lambda r: (str(r["sample"]), int(r["injection_idx"])) in drop_pairs, axis=1)
        res.df = res.df[~mask].copy()
    return res


@st.fragment
def _exclusion_ui_fragment(eid, folder, group_labels):
    """剔除页主体。fragment：勾选/编辑只重渲染本区块。"""
    _exclusion_ui(eid, folder, group_labels)


def _exclusion_ui(eid, folder, group_labels):
    conn = get_db()
    folder_path = config.EXPERIMENTS_DIR / folder
    # 解析缓存：勾选触发 rerun 时不再重读 txt（解决卡顿）；缓存的是套用了标准品剔除的计算用结果
    cache = st.session_state.setdefault("_excl_res_cache", {})
    cache_sig = st.session_state.setdefault("_excl_res_cache_sig", {})
    if folder not in cache or cache_sig.get(folder) != _folder_sig(folder):
        try:
            cache[folder] = _load_calc_res(conn, eid, folder)
            cache_sig[folder] = _folder_sig(folder)
        except Exception:
            st.warning("无法重新解析原始数据。请确认实验文件夹的 txt 仍在。")
            return
    res = cache[folder]
    _saved_excl_ui = M.get_excluded_samples(conn, eid)

    st.markdown("**Excel 式剔除**：在表格「剔除」列勾选（该列与样本名一同冻结，左右滚动不消失）→ "
                "下方**实时重算**剔除后平均值/RSD（红=剔除后RSD仍>15%）；确认无误后点底部「保存」。")
    n_rep = res.n_replicates
    # 预聚合（一次 groupby 取代此前 组×相×物质×重复 的 ~2500 次全表布尔扫描）：
    # _cell[(样本名, 物质)] = 峰面积；_gpv[(组,相,物质)] = [(样本名, 面积), ...]
    _smp_df = res.df[res.df["category"] == "sample"]
    _cell = {}
    for _s, _a, _v in zip(_smp_df["sample"], _smp_df["analyte"], _smp_df["area"]):
        _cell.setdefault((_s, _a), float(_v))
    _gpv = {}
    for (_g_, _ph_, _a_), _sub in _smp_df.groupby(["group", "phase", "analyte"]):
        try:
            _gpv[(int(_g_), str(_ph_), _a_)] = list(zip(_sub["sample"], _sub["area"]))
        except Exception:
            continue
    # (组, 相, 重复) → 真实样品名（样品名格式可自定义，不能拼接）
    _name_map = {}
    for _s_, _g_, _ph_, _r_ in zip(_smp_df["sample"], _smp_df["group"], _smp_df["phase"], _smp_df["replicate"]):
        try:
            _name_map[(int(_g_), str(_ph_), int(_r_))] = _s_
        except Exception:
            continue

    def _stat_pairs(pairs, excl_names):
        """[(样本名, 面积)] → (均值, RSD%)，语义与 calculator.mean/rsd_percent 一致"""
        vals = [float(v) for s_, v in pairs if s_ not in excl_names]
        if not vals:
            return 0.0, 0.0
        m = float(np.mean(vals))
        r = 0.0
        if len(vals) >= 2 and m != 0:
            r = float(np.std(vals, ddof=1) / m * 100.0)
        return m, r

    all_excl = []
    for g in res.groups:
        lbl = group_labels.get(g, f"M{g}")
        with st.expander(f"M{g}（{lbl}）", expanded=False):
            for phase, phase_name in [("A", "受体相"), ("S", "供体相")]:
                # 检查该组该相是否有数据
                if not any((g, phase, a_) in _gpv for a_ in res.analytes):
                    continue
                st.markdown(f"**{phase_name}**（「剔除」列与样本名一同冻结，左右滚动查看各物质峰面积）")
                # 合并表：样本名（索引，冻结）+ 剔除复选框（冻结）+ 各物质峰面积（只读列）
                _ed_rows = []
                for r in range(1, n_rep + 1):
                    sname = _name_map.get((g, phase, r), f"{g}-{phase}-{r}")
                    _row_e = {"样本": sname, "剔除": sname in _saved_excl_ui}
                    for _a_ in res.analytes:
                        _v = _cell.get((sname, _a_))
                        _row_e[_a_] = float(_v) if _v is not None else np.nan
                    _ed_rows.append(_row_e)
                _ed_w = pd.DataFrame(_ed_rows).set_index("样本")
                _area_cols = list(res.analytes)
                _ed_t = st.data_editor(
                    _ed_w,
                    key=f"ex3_{eid}_{g}_{phase}",
                    column_config={
                        "剔除": st.column_config.CheckboxColumn("剔除", default=False, pinned=True),
                        **{c: st.column_config.NumberColumn(c, format="%.2e") for c in _area_cols}},
                    disabled=_area_cols,
                    width="stretch",
                    hide_index=False)
                _orig_names = set(_ed_w.index)
                excl_here = [s for s in _ed_t.index[_ed_t["剔除"]] if s in _orig_names]  # 防误改索引文字
                excl_set = set(excl_here)
                # 实时重算统计表（行=物质，窄表无横向滚动）
                rows = []
                for analyte in res.analytes:
                    pairs = _gpv.get((g, phase, analyte)) or []
                    m_raw, rsd_raw = _stat_pairs(pairs, ())
                    m_fix, rsd_fix = _stat_pairs(pairs, excl_set)
                    rows.append({"物质": analyte,
                                 "剔除后平均值": m_fix, "剔除后RSD%": round(rsd_fix, 2),
                                 "原始平均值": m_raw, "原始RSD%": round(rsd_raw, 2)})
                tbl = pd.DataFrame(rows).set_index("物质")

                def _style_row(s):
                    styles = []
                    for c in s.index:
                        if c in ("剔除后RSD%", "原始RSD%") and pd.notna(s[c]) and float(s[c]) > 15.0:
                            styles.append("background-color: #ffc7ce; color: #9C0006")
                        else:
                            styles.append("")
                    return styles
                # 实时统计置于 RSD 表之前：勾选即可见总数变化，无需滚动
                _n_over = int((tbl["剔除后RSD%"] > 15).sum())
                _rsd_sum = float(tbl["剔除后RSD%"].sum())
                st.caption(f"📊 实时统计：剔除后 RSD>15% 的物质 **{_n_over}** 个 | "
                           f"所有物质剔除后 RSD% 累加 **{_rsd_sum:.1f}**")
                # 物质名列（索引）必须显示——否则统计表没有标题列
                _tbl_show = tbl.rename_axis("物质").reset_index()
                st.dataframe(_tbl_show.style.apply(_style_row, axis=1)
                             .format({"剔除后平均值": "{:.2e}", "原始平均值": "{:.2e}",
                                      "剔除后RSD%": "{:.2f}", "原始RSD%": "{:.2f}"}, na_rep="-"),
                             width="stretch", hide_index=True)
                all_excl.extend(excl_here)

    all_excl = list(dict.fromkeys(all_excl))
    all_excl = [s for s in all_excl if s in set(_smp_df["sample"])]  # 滤掉缺数据重复位的回退名（幻影行）
    st.markdown("---")
    st.markdown(f"**已勾选剔除**（{len(all_excl)} 个）：{', '.join(all_excl) if all_excl else '无'}")
    if all_excl and st.button("💾 确认剔除并保存到数据库 + DATA.xlsx", type="primary", key=f"save_{eid}"):
        # 用该实验自己的 Va/Vd/启用供体相 设置重算（不能用全局默认值），并保留标准品剔除状态
        exp_date_row = conn.execute("SELECT exp_date, Va, Vd, use_sample, data_xlsx, temperature, humidity, column_pressure, purpose FROM experiments WHERE id=?", (eid,)).fetchone()
        va_e = float(exp_date_row["Va"]) if exp_date_row and exp_date_row["Va"] else float(config.DEFAULT_VA)
        vd_e = float(exp_date_row["Vd"]) if exp_date_row and exp_date_row["Vd"] else float(config.DEFAULT_VD)
        us_e = bool(exp_date_row["use_sample"]) if exp_date_row else True
        exp_date = exp_date_row["exp_date"] if exp_date_row else ""
        _gqc_ex, _dqc_ex, _ = _build_group_qc_from_db(conn, eid)
        m2 = C.compute_metrics(res.df, res.analytes, res.groups,
                               va_e, vd_e,
                               use_sample=us_e, excluded_samples=set(all_excl),
                               group_qc=_gqc_ex, default_qc=_dqc_ex)
        conn.execute("DELETE FROM metrics WHERE experiment_id=?", (eid,))
        conn.commit()
        M.save_metrics(conn, eid, exp_date, m2, res.analytes, res.groups)
        M.set_excluded_samples(conn, eid, set(all_excl))  # 持久化剔除集（有效膜数等统计用）
        fname = ""
        try:
            va = float(exp_date_row["Va"]) if exp_date_row["Va"] else float(config.DEFAULT_VA)
            vd = float(exp_date_row["Vd"]) if exp_date_row["Vd"] else float(config.DEFAULT_VD)
            us = bool(exp_date_row["use_sample"]) if exp_date_row else True
            data_xlsx_path = Path(exp_date_row["data_xlsx"]) if exp_date_row and exp_date_row["data_xlsx"] else None
            if data_xlsx_path:
                _eid_row = conn.execute("SELECT current_sheets FROM experiments WHERE id=?", (eid,)).fetchone()
                _cur_maps_x = json.loads(_eid_row["current_sheets"]) if _eid_row and _eid_row["current_sheets"] else []
                _gqc_bx, _dqc_bx, _ = _build_group_qc_from_db(conn, eid)
                # 完整元数据：薄 meta 会让重建的总表把温度/湿度/柱压/实验说明清成空白
                _meta_bx = {"exp_date": exp_date,
                            "temperature": exp_date_row["temperature"], "humidity": exp_date_row["humidity"],
                            "column_pressure": exp_date_row["column_pressure"], "purpose": exp_date_row["purpose"]}
                EB.build_excel(data_xlsx_path, res, _meta_bx, va, vd, us,
                               excluded_samples=set(all_excl),
                               gcond=M.get_group_conditions(conn, eid),
                               voltage_df=build_voltage_df(_cur_maps_x),
                               qc_by_group=_gqc_bx, default_qc=_dqc_bx)
                fname = data_xlsx_path.name
        except Exception as e:
            st.caption(f"（Excel 同步更新失败：{e}）")
        # 一次性提示：先存 session_state 再 rerun（直接 st.success 后 rerun 用户看不到）
        st.session_state[f"_excl_saved_msg_{eid}"] = (
            f"✅ 已剔除 {len(all_excl)} 个样本，**已保存到数据库并重新生成 DATA.xlsx**"
            + (f"（`{fname}`）" if fname else "")
            + "。热图 tab 已同步刷新。")
        # 清缓存使热图立即反映新数据
        st.session_state["_excl_res_cache"].pop(folder, None)
        st.rerun()
    # rerun 后显示一次性提示（显示即清除）
    if st.session_state.get(f"_excl_saved_msg_{eid}"):
        st.success(st.session_state[f"_excl_saved_msg_{eid}"])
        del st.session_state[f"_excl_saved_msg_{eid}"]


@st.fragment
def _edit_experiment_fragment(eid, row, folder):
    """修改/删除页主体。fragment：页内交互（元数据/组合/剔除矩阵编辑）只重渲染本区块，
    不再整页执行全部 6 个标签页（热图/三相分布/剔除等），交互明显提速。"""
    _edit_experiment_ui(eid, row, folder)


def _edit_experiment_ui(eid, row, folder):
    conn = get_db()
    st.caption("修改元数据或组别名后，点「重新生成」会刷新 DATA.xlsx 和数据库。")
    ec = st.columns(2)
    new_temp = ec[0].text_input("温度(℃)", value=row["temperature"] or "", key=f"ed_temp_{eid}")
    new_hum = ec[1].text_input("湿度(%)", value=row["humidity"] or "", key=f"ed_hum_{eid}")
    new_cp = st.text_input("液相平衡柱压", value=row["column_pressure"] or "", key=f"ed_cp_{eid}")
    ec2 = st.columns(2)
    new_va = ec2[0].number_input("Va(μL)", value=float(row["Va"]), key=f"ed_va_{eid}")
    new_vd = ec2[1].number_input("Vd(μL)", value=float(row["Vd"]), key=f"ed_vd_{eid}")
    new_purpose = st.text_area("实验说明", value=row["purpose"] or "", key=f"ed_purpose_{eid}")

    # —— 多标准品组合（编辑：改组合定义 + 各组选择）——
    # 组合列表存会话键（与新增实验页同模式）：➕添加按钮写入后下一轮能读到，
    # 否则每次渲染都从数据库重读、添加的组合立刻丢失
    _combo_state_key = f"ed_{eid}_qc_combos"
    try:
        _combos_ed = json.loads(row["qc_combos"]) if row["qc_combos"] else []
    except Exception:
        _combos_ed = []
    st.markdown("---")
    st.markdown("**🧪 标准品组合**（多标准品模式：定义多套供受体相质控组合，逐组选择）")
    _multi_ed = st.checkbox("实验组使用不同的供受体相质控样", value=bool(_combos_ed),
                            key=f"ne_multi_qc_ed_{eid}")
    if not _multi_ed:
        st.session_state.pop(_combo_state_key, None)
    _combos_new_ed = _combos_ed
    if _multi_ed:
        # 候选 = 解析出的全部样品名（含各组合的标准品/空白）+ DB 质控名
        try:
            _base_qn_ed = json.loads(row["qc_names"]) if row["qc_names"] else {}
            _base_qn_ed = {k: (v[0] if isinstance(v, list) and v else v) for k, v in _base_qn_ed.items()}
        except Exception:
            _base_qn_ed = {}
        try:
            _res_names_ed = _cached_parse(folder, "", _row_fmt := (row["sample_format"] if "sample_format" in row.keys() else None) or "",
                                          _folder_sig(folder)).all_sample_names
        except Exception:
            _res_names_ed = []
        _all_samples_ed = sorted({v for v in (list(_res_names_ed) + list(_base_qn_ed.values())) if isinstance(v, str) and v})
        if _combo_state_key not in st.session_state:
            st.session_state[_combo_state_key] = _combos_ed or [
                {"name": "默认组合", **{k: _base_qn_ed.get(k, "") for k, _ in QC_ROLES}}]
        _combos_new_ed = _render_qc_combo_editor(_all_samples_ed, st.session_state[_combo_state_key], f"ed_{eid}")
        st.session_state[_combo_state_key] = _combos_new_ed
    # 每组选择（在下方组条件编辑器里）
    _combos_for_groups_ed = _combos_new_ed if _multi_ed else []

    gcond = M.get_group_conditions(conn, eid)
    new_labels = {}
    new_group_conds = {}  # {组号: {字段: 新值}}
    gc_old_pass = {}  # {组号: {已从界面移除的字段: 旧值}}——保存时透传，不覆盖历史
    if len(gcond):
        st.markdown("**组别名 + 组级条件编辑**")
        # 可编辑的组级字段
        gc_fields = [("slm_main", "主SLM"), ("pH_donor", "供体相pH"), ("donor_composition", "供体相成分"),
                     ("donor_buffer_conc", "供体相缓冲剂浓度"),
                     ("pH_acceptor", "受体相pH"), ("acceptor_composition", "受体相成分"),
                     ("acceptor_buffer_conc", "受体相缓冲剂浓度"),
                     ("extract_time", "萃取时间"), ("extract_voltage", "萃取电压")]
        for _, gr in gcond.iterrows():
            if gr["grp"] == 0:
                continue
            gnum = int(gr["grp"])
            with st.expander(f"M{gnum}"):
                new_labels[gnum] = st.text_input(
                    "组别名", value=gr["group_label"] or f"M{gnum}", key=f"ed_lbl_{eid}_{gnum}")
                gc_new = {}
                # 标准品组合选择（多标准品模式）
                if _multi_ed and _combos_for_groups_ed:
                    _cn_opts_ed = ["（默认组合）"] + [c.get("name") for c in _combos_for_groups_ed]
                    gc_new["qc_combo"] = st.selectbox(
                        f"M{gnum} 标准品组合（供受体相质控）", _cn_opts_ed,
                        index=(_cn_opts_ed.index(gr["qc_combo"]) if gr.get("qc_combo") in _cn_opts_ed else 0),
                        key=f"ed_{eid}_{gnum}_qc_combo")
                    if gc_new["qc_combo"] == "（默认组合）":
                        gc_new["qc_combo"] = ""
                else:
                    gc_new["qc_combo"] = ""
                gcc = st.columns(3)
                for i, (fk, fl) in enumerate(gc_fields):
                    col = gcc[i % 3]
                    cur_val = gr[fk] if fk in gcond.columns else ""
                    if fk in ("pH_donor", "pH_acceptor", "extract_time", "extract_voltage"):
                        try:
                            cur_num = float(cur_val) if cur_val not in (None, "") else 0.0
                        except (ValueError, TypeError):
                            cur_num = 0.0
                        gc_new[fk] = col.number_input(fl, value=cur_num, key=f"ed_{eid}_{gnum}_{fk}")
                    else:
                        gc_new[fk] = col.text_input(fl, value=str(cur_val or ""), key=f"ed_{eid}_{gnum}_{fk}")
                # 异常标签 + 孔位（多选）+ 异常孔位
                try:
                    _cur_tags = json.loads(gr["note_tags"]) if gr.get("note_tags") else []
                except Exception:
                    _cur_tags = []
                edc2 = st.columns(3)
                # 选项 = 预设标签 + 已保存的自定义标签（默认值必须在选项里，否则整页崩溃）
                _tag_opts = list(dict.fromkeys(list(S.get_note_tags()) + [t for t in _cur_tags if t]))
                gc_new["note_tags"] = edc2[0].multiselect(
                    "异常标签", _tag_opts, default=[t for t in _cur_tags if t in _tag_opts],
                    key=f"ed_{eid}_{gnum}_tags", accept_new_options=True, placeholder="渗漏/烧焦/破裂…")
                _well_list_ed = [f"{r}{c}" for r in "ABCDEFGH" for c in range(1, 13)]
                _cur_wells = M.parse_wells(gr["well_position"] if gr.get("well_position") else "")
                gc_new["well_position"] = edc2[1].multiselect(
                    "孔位(可多选)", _well_list_ed, default=[w for w in _cur_wells if w in _well_list_ed],
                    key=f"ed_{eid}_{gnum}_well", accept_new_options=True, placeholder="该组使用的孔")
                _cur_abw = M.parse_wells(gr["abnormal_wells"] if gr.get("abnormal_wells") else "")
                # 联动：已填孔位 → 只显示那些孔；未填 → 全部96孔
                _abw_opts_ed = _cur_wells or _well_list_ed
                gc_new["abnormal_wells"] = edc2[2].multiselect(
                    "异常孔位(可选)", _abw_opts_ed,
                    default=[w for w in _cur_abw if w in _abw_opts_ed],
                    key=f"ed_{eid}_{gnum}_abw", help="空=该组全部孔位视为异常")
                gc_old_pass[gnum] = {"donor_additive_conc": gr["donor_additive_conc"],
                                     "acceptor_additive_conc": gr["acceptor_additive_conc"]}
                new_group_conds[gnum] = gc_new

    # —— 电流记录重新选择（可修正录入时选错/漏选）——
    st.markdown("**⚡ 电流记录重新选择**（修正录入时选错或漏选）")
    cur_up_ed = st.file_uploader("上传新电流 xlsx（保存到 data/current/）", type=["xlsx"],
                                 key=f"ed_cur_up_{eid}")
    if cur_up_ed:
        if st.button("💾 保存上传的电流文件", key=f"ed_cur_save_{eid}"):
            CUR.CURRENT_DIR.mkdir(parents=True, exist_ok=True)
            (CUR.CURRENT_DIR / cur_up_ed.name).write_bytes(cur_up_ed.getvalue())
            st.success(f"已保存 `{cur_up_ed.name}`")
            st.rerun()
    cur_sheet_list_ed = CUR.list_all_sheets_detailed()
    if cur_sheet_list_ed:
        _opts_ed = [f"{d['file']} :: {d['sheet']}" for d in cur_sheet_list_ed]
        # 预选当前已保存的映射
        _cur_saved = []
        try:
            _cur_saved = [f"{m['file']} :: {m['sheet']}" for m in json.loads(row["current_sheets"])] if row["current_sheets"] else []
        except Exception:
            _cur_saved = []
        sel_sheets_ed = st.multiselect("选择本次实验的电流 sheet（可多选；空=清除映射，按命名规则自动匹配）",
                                       _opts_ed, default=[o for o in _cur_saved if o in _opts_ed],
                                       key=f"ed_cur_sel_{eid}")
        st.session_state[f"ed_current_sheets_{eid}"] = [
            {"file": s.split(" :: ")[0], "sheet": s.split(" :: ", 1)[1]} for s in sel_sheets_ed]
    else:
        st.caption("`data/current/` 为空，可上传 xlsx。")

    # —— 标准品剔除（修改时也可用；实时预览，与新增实验一致）——
    st.markdown("---")
    st.markdown("**🧪 标准品剔除重算**（受体相/供体相两套矩阵分别勾选，RSD 实时更新 → 底部「保存修改并重新生成」一并保存）")
    folder_path = config.EXPERIMENTS_DIR / folder
    _std_link_ed = st.checkbox(
        "同一标准品名在两相都使用时联动剔除（勾选 A-STD 第 i 针，另一相的 A-STD 第 i 针同剔；不同名标准品互不影响）",
        value=True, key=f"std_link_ed_{eid}",
        help="两相用同一标准品系列（同名或同批次对应针）时勾选；只需在一处勾选，两相同号针一起剔除。")
    excl_by_phase = {}
    _std_names_by_phase_ed = {}
    std_ui_ok = False
    try:
        try:
            _qc_names_prev = json.loads(row["qc_names"]) if row["qc_names"] else {}
        except Exception:
            _qc_names_prev = {}
        try:
            _combos_ed_parse = json.loads(row["qc_combos"]) if row["qc_combos"] else []
        except Exception:
            _combos_ed_parse = []
        if _combos_ed_parse:
            _qc_names_prev = _combos_union_qc(_combos_ed_parse)
        _row_fmt_ed = row["sample_format"] if "sample_format" in row.keys() else None
        res_edit = _cached_parse(folder, json.dumps(_qc_names_prev, ensure_ascii=False) if _qc_names_prev else "",
                                 _row_fmt_ed or "", _folder_sig(folder))
        # 已保存的标准品针剔除（在勾选框中预勾选，取消勾选即恢复该针）
        _std_excl_saved = M.get_excluded_std(conn, eid)
        _saved_needles = {f"{s} #{i}" for pairs in _std_excl_saved.values() for s, i in pairs}

        def _render_std_editor(mat, keyp):
            row_sum = mat.sum(axis=1)
            mean_sum = row_sum.mean()
            dev = (row_sum - mean_sum).abs() / (abs(mean_sum) if mean_sum else 1)
            worst = dev.idxmax()
            st.warning(f"🔴 推荐剔除（偏离最大）：**{worst}**（偏离度 {dev.max():.3f}）")
            # 合并表：进样针（索引，冻结）+ 剔除复选框（冻结）+ 峰面积总和/偏离度/各物质（只读）
            _m_ed = mat.copy()
            _m_ed.insert(0, "峰面积总和", row_sum)
            _m_ed.insert(1, "偏离度", dev.round(3))
            _m_ed.insert(0, "剔除", [n in _saved_needles for n in mat.index])
            _num_cols_ed = [c for c in _m_ed.columns if c != "剔除"]
            _ed_m = st.data_editor(
                _m_ed,
                key=f"ed_std_ed3_{keyp}_{eid}", hide_index=False,
                column_config={
                    "剔除": st.column_config.CheckboxColumn("剔除", default=False, pinned=True),
                    "峰面积总和": st.column_config.NumberColumn(format="%.3e"),
                    "偏离度": st.column_config.NumberColumn(format="%.3f"),
                    **{c: st.column_config.NumberColumn(format="%.3e") for c in mat.columns}},
                disabled=_num_cols_ed)
            return [n for n in _ed_m.index[_ed_m["剔除"]].tolist() if n in mat.index]  # 防误改索引文字

        def _phase_rsd(df_, cat, analytes_):
            out_ = {}
            for a_ in analytes_:
                _, rv, _ = C.aggregate_qc(df_, cat, a_)
                out_[a_] = round(rv, 2) if rv is not None else None
            return out_

        def _hl_ed2(v):
            if pd.isna(v): return ""
            try:
                return "background-color: #ffc7ce; color: #9C0006" if float(v) > 15.0 else ""
            except (ValueError, TypeError):
                return ""

        excl_by_phase = {}
        for phase_cat, phase_name in [("std_acceptor", "受体相标准品"), ("std_donor", "供体相标准品")]:
            _std_names_by_phase_ed[phase_cat] = sorted(
                res_edit.df[res_edit.df["category"] == phase_cat]["sample"].unique())
            if not _std_names_by_phase_ed[phase_cat]:
                st.caption(f"（{phase_name}：数据中无对应进样行）")
                continue
            st.markdown(f"**{phase_name}**（按标准品名分节剔除；剔除某针后所有使用该标准品的组同时生效）")
            excl_by_phase.setdefault(phase_cat, [])
            # —— 每个标准品名独立一节（不同溶液峰面积不可混算）——
            for _nm in _std_names_by_phase_ed[phase_cat]:
                _df_nm = res_edit.df[(res_edit.df["category"] == phase_cat)
                                     & (res_edit.df["sample"] == _nm)]
                std_mat_ed = C.get_std_injection_matrix(_df_nm, phase_cat, res_edit.analytes)
                if not len(std_mat_ed):
                    continue
                with st.expander(f"📌 {_nm}（{len(std_mat_ed)} 针）",
                                 expanded=(_nm == _std_names_by_phase_ed[phase_cat][0])):
                    excl_by_phase[phase_cat].extend(_render_std_editor(std_mat_ed, f"{phase_cat}_{_nm}"))

        # 合并两相剔除对（同系列联动：按针号对应扩展到另一相）
        # 注意：std_ui_ok 必须在 _pairs_by_phase_ed 构建成功后才置 True，
        # 否则中途异常（如 data_editor 动态行产生空值）会让保存按钮引用未定义变量
        _pairs_raw_ed = {}
        for _pc, _needles in excl_by_phase.items():
            _pp = []
            for needle in _needles:
                if isinstance(needle, str) and " #" in needle:
                    sn_, n_ = needle.rsplit(" #", 1)
                    _pp.append((sn_, int(n_)))
            _pairs_raw_ed[_pc] = _pp
        _pairs_by_phase_ed = _expand_std_link(_pairs_raw_ed, _std_names_by_phase_ed) if _std_link_ed else _pairs_raw_ed
        excl_pairs_all = [(s, i) for pairs in _pairs_by_phase_ed.values() for s, i in pairs]
        std_ui_ok = True

        # 实时 RSD 预览：按标准品名分别显示剔除前/后 RSD（不再跨溶液混算）
        st.markdown("**各物质·各标准品 RSD（勾选剔除后实时更新）**（红色=RSD>15%）")
        _excl_set_ed = set(excl_pairs_all)
        for phase_cat, phase_label in [("std_acceptor", "受体"), ("std_donor", "供体")]:
            for _nm in _std_names_by_phase_ed.get(phase_cat, []):
                _excl_nm = {(s, i) for (s, i) in _excl_set_ed if s == _nm}
                _df_b = res_edit.df[(res_edit.df["category"] == phase_cat)
                                    & (res_edit.df["sample"] == _nm)]
                if not len(_df_b):
                    continue
                _df_a = _df_b[~_df_b.apply(
                    lambda r: (r["sample"], int(r["injection_idx"])) in _excl_nm, axis=1)]
                _col_b, _col_a = f"{_nm} 剔除前RSD%", f"{_nm} 剔除后RSD%"
                _rows_pn = []
                for a in res_edit.analytes:
                    _, r_b, _ = C.aggregate_qc(_df_b, phase_cat, a)
                    _, r_a, _ = C.aggregate_qc(_df_a, phase_cat, a)
                    _rows_pn.append({"物质": a,
                                     _col_b: round(r_b, 2) if r_b is not None else None,
                                     _col_a: round(r_a, 2) if r_a is not None else None})
                _rdf = pd.DataFrame(_rows_pn)
                _rdf[f"{_nm} 变化"] = (_rdf[_col_a] - _rdf[_col_b]).round(2)
                with st.expander(f"{phase_label}相 · {_nm}（{len(_rdf)} 个物质）",
                                 expanded=len(_std_names_by_phase_ed.get(phase_cat, [])) <= 2):
                    _cols_pn = [c for c in _rdf.columns if c != "物质"]
                    # 实时统计置于 RSD 表之前：勾选即可见总数变化
                    _n_over = int((_rdf[_col_a] > 15).sum())
                    _rsd_sum = float(_rdf[_col_a].sum())
                    st.caption(f"📊 实时统计：剔除后 RSD>15% 的物质 **{_n_over}** 个 | "
                               f"所有物质剔除后 RSD% 累加 **{_rsd_sum:.1f}**")
                    st.dataframe(_rdf.style.map(_hl_ed2, subset=_cols_pn).format(
                        "{:.2f}", na_rep="-", subset=_cols_pn),
                        width="stretch", hide_index=True)

        # 保存结果一次性提示（按钮在下方：保存修改并重新生成 DATA.xlsx）
        msg_key_ed2 = f"_ed_regen_msg_{eid}"
        if st.session_state.get(msg_key_ed2):
            st.success(st.session_state[msg_key_ed2])
            del st.session_state[msg_key_ed2]

    except Exception as e:
        st.caption(f"（标准品数据加载失败：{e}）")

    if st.button("💾 保存修改并重新生成 DATA.xlsx", type="primary", key=f"ed_regen_{eid}"):
        conn.execute(
            "UPDATE experiments SET temperature=?, humidity=?, column_pressure=?, Va=?, Vd=?, purpose=?, current_sheets=? WHERE id=?",
            (new_temp, new_hum, new_cp, new_va, new_vd, new_purpose,
             json.dumps(st.session_state.get(f"ed_current_sheets_{eid}") or [], ensure_ascii=False), eid))
        if len(gcond):
            # 更新组别名 + 组级条件（含异常标签/孔位）
            for grp, gc_new in new_group_conds.items():
                set_clause = ("group_label=?, slm_main=?, pH_donor=?, donor_composition=?, donor_buffer_conc=?, "
                              "donor_additive_conc=?, pH_acceptor=?, acceptor_composition=?, acceptor_buffer_conc=?, "
                              "acceptor_additive_conc=?, extract_time=?, extract_voltage=?, note_tags=?, well_position=?, abnormal_wells=?, qc_combo=?")
                vals = [new_labels.get(grp, f"M{grp}"),
                        gc_new.get("slm_main"), gc_new.get("pH_donor"), gc_new.get("donor_composition"),
                        gc_new.get("donor_buffer_conc"), gc_old_pass.get(grp, {}).get("donor_additive_conc"),
                        gc_new.get("pH_acceptor"), gc_new.get("acceptor_composition"),
                        gc_new.get("acceptor_buffer_conc"), gc_old_pass.get(grp, {}).get("acceptor_additive_conc"),
                        gc_new.get("extract_time"), gc_new.get("extract_voltage"),
                        json.dumps(gc_new.get("note_tags") or [], ensure_ascii=False),
                        json.dumps(gc_new.get("well_position") or [], ensure_ascii=False),
                        json.dumps(gc_new.get("abnormal_wells") or [], ensure_ascii=False),
                        gc_new.get("qc_combo") or ""]
                conn.execute(f"UPDATE group_conditions SET {set_clause} WHERE experiment_id=? AND grp=?",
                             tuple(vals) + (eid, grp))
        conn.commit()
        # 标准品针剔除：随本次保存一并落库（联动扩展后按相存储；UI 未加载成功则保持原状）
        # 标准品组合配置随保存落库
        _combos_final_ed = st.session_state.get(f"ed_{eid}_qc_combos", _combos_new_ed) if _multi_ed else None
        conn.execute("UPDATE experiments SET qc_combos=? WHERE id=?",
                     (json.dumps(_combos_final_ed, ensure_ascii=False) if _multi_ed and _combos_final_ed else None, eid))
        conn.commit()
        if std_ui_ok:
            # 与步骤6 同语义：本次矩阵中出现并勾选=保留；出现未勾选=恢复；矩阵未出现的旧针=保留
            # （防止历史剔除的标准品名从本次解析中消失时被全量覆盖清空）
            _final_std_ed = {}
            for _pc_ed in ("std_acceptor", "std_donor"):
                _new_ed = set(_pairs_by_phase_ed.get(_pc_ed) or [])
                _old_ed = set(_std_excl_saved.get(_pc_ed) or set())
                _seen_ed = {(s, i) for s, i in _old_ed if s in _std_names_by_phase_ed.get(_pc_ed, set())}
                _final_std_ed[_pc_ed] = list(_new_ed | (_old_ed - _seen_ed))
            M.set_excluded_std(conn, eid,
                               std_acceptor=_final_std_ed.get("std_acceptor"),
                               std_donor=_final_std_ed.get("std_donor"))
        try:
            res = _load_calc_res(conn, eid, folder)  # 带 qc_names + 刚保存的标准品针剔除
            _saved_sample_excl3 = M.get_excluded_samples(conn, eid)
            _gqc_ed, _dqc_ed, _ = _build_group_qc_from_db(conn, eid)
            metrics = C.compute_metrics(res.df, res.analytes, res.groups, new_va, new_vd,
                                        use_sample=bool(row["use_sample"]),
                                        excluded_samples=_saved_sample_excl3,
                                        group_qc=_gqc_ed, default_qc=_dqc_ed)
            data_xlsx = folder_path / f"{row['exp_date'] or res.file_date}-DATA.xlsx"
            meta = {"exp_date": row["exp_date"],
                    "temperature": new_temp, "humidity": new_hum, "column_pressure": new_cp,
                    "Va": new_va, "Vd": new_vd, "purpose": new_purpose}
            # 同步：电流记录 + 组条件（含备注/标签/孔位）一并写入 DATA.xlsx
            # （用刚保存的选择，而不是 row 里本轮更新前的旧值）
            _cur_maps = st.session_state.get(f"ed_current_sheets_{eid}") or []
            EB.build_excel(data_xlsx, res, meta, new_va, new_vd, bool(row["use_sample"]),
                           voltage_df=build_voltage_df(_cur_maps),
                           gcond=M.get_group_conditions(conn, eid),
                           excluded_samples=_saved_sample_excl3,
                           qc_by_group=_gqc_ed, default_qc=_dqc_ed)
            conn.execute("DELETE FROM metrics WHERE experiment_id=?", (eid,))
            conn.execute("DELETE FROM qc_summary WHERE experiment_id=?", (eid,))
            M.save_metrics(conn, eid, row["exp_date"], metrics, res.analytes, res.groups)
            M.save_qc_summary(conn, eid, row["exp_date"], res, res.analytes, default_qc=_dqc_ed)
            st.session_state.get("_excl_res_cache", {}).pop(folder, None)
            _n_acc_m = len(_pairs_by_phase_ed.get("std_acceptor") or []) if std_ui_ok else 0
            _n_don_m = len(_pairs_by_phase_ed.get("std_donor") or []) if std_ui_ok else 0
            st.session_state[f"_ed_regen_msg_{eid}"] = (
                f"✅ 已保存并重新生成 `{data_xlsx.name}`：元数据/组条件已更新，"
                f"标准品剔除（受体相 {_n_acc_m} 针、供体相 {_n_don_m} 针，含联动）已应用，"
                f"均值/RSD/回收率已重算。")
            st.rerun()
        except Exception as e:
            conn.rollback()
            st.error(f"重新生成失败：{e}")


# ========================================================
# 3. 质控看板
# ========================================================
def page_qc_dashboard():
    st.title("📊 质控看板")
    conn = get_db()
    qc = M.get_qc_df(conn)
    if len(qc) == 0:
        st.info("暂无质控数据。"); return

    render_detail_sidebar()
    _qc_dashboard_fragment(qc, conn)


@st.fragment
def _qc_dashboard_fragment(qc: pd.DataFrame, conn):
    """质控看板主体（fragment 局部刷新：切换分区/改筛选只重跑本区块，不再整页重渲染，解决卡顿）"""
    qsel = st.radio("质控看板分区", ["跨实验-峰面积质控", "跨实验-保留时间/半峰宽质控", "单次实验RSD柱状图",
                                    "单次实验保留时间散点", "🌡️ 实验环境质控", "📈 单次实验半峰宽散点",
                                    "📈 跨实验RSD超阈统计", "⚠️ 异常统计(渗漏/烧焦/破裂)"],
                    horizontal=True, key="qc_tab_sel")

    # ---- 跨实验质控（物质可多选，叠加在同一张图用不同颜色）----
    analytes = sorted(qc["analyte"].unique())
    if qsel == "跨实验-峰面积质控":
        a_sel = st.multiselect("选择物质（可多选，叠加同图）", analytes,
                               default=[analytes[0]], key="qc_area_a")
        st.plotly_chart(D.qc_area_chart(qc, a_sel or analytes[:1]), width="stretch",
                        key="qc_area_overlay")
    if qsel == "跨实验-保留时间/半峰宽质控":
        a_sel2 = st.multiselect("选择物质（可多选，叠加同图）", analytes,
                                default=[analytes[0]], key="qc_rt_a")
        st.plotly_chart(D.qc_rt_chart(qc, a_sel2 or analytes[:1]), width="stretch",
                        key="qc_rt_overlay")
        # 半峰宽跨实验质控（与RT同类）
        if "std_width_mean" in qc.columns and qc["std_width_mean"].notna().any():
            st.markdown("---")
            st.markdown("**跨实验-半峰宽质控**（标准品半峰宽均值 vs 日期，各物质 ±15% 点线参考带）")
            st.plotly_chart(D.qc_width_chart(qc, a_sel2 or analytes[:1]), width="stretch",
                            key="qc_width_overlay")

    # ---- 单次实验 RSD 柱状图（选日期，X=分析物）----
    if qsel == "单次实验RSD柱状图":
        dates = sorted(qc["exp_date"].dropna().unique())
        if not dates:
            st.caption("无数据"); 
        else:
            sel_date = st.selectbox("选择实验日期", dates, key="qc_rsd_date")
            sub = qc[qc["exp_date"] == sel_date].sort_values("analyte")
            if len(sub):
                fig = go.Figure()
                fig.add_trace(go.Bar(
                    x=sub["analyte"], y=sub["std_acceptor_rsd"],
                    marker_color=["#9C0006" if (pd.notna(v) and v > 15) else "#4472C4" for v in sub["std_acceptor_rsd"]],
                    text=[f"{v:.1f}%" if pd.notna(v) else "" for v in sub["std_acceptor_rsd"]],
                    textposition="outside"))
                fig.add_hline(y=15, line_dash="dash", line_color="red", annotation_text="15%阈值")
                fig.update_layout(title=f"实验 {sel_date} 各分析物标准品 RSD%",
                                  xaxis_title="分析物", yaxis_title="受体标准品RSD%",
                                  template="simple_white", height=500)
                st.plotly_chart(fig, width="stretch")

    # ---- 单次实验保留时间散点（默认单物质）+ RT波动 ----
    if qsel == "单次实验保留时间散点":
        rt_df = M.get_rt_df(conn)
        rt_dates = sorted(rt_df["exp_date"].dropna().unique()) if len(rt_df) else []
        if not rt_dates:
            st.caption("无RT数据")
        else:
            rd1, rd2 = st.columns(2)
            sel_rt_date = rd1.selectbox("选择实验日期", rt_dates, key="rt_scatter_date")
            rt_analytes = sorted(rt_df[rt_df["exp_date"] == sel_rt_date]["analyte"].unique())
            sel_rt_a = rd2.selectbox("选择物质", ["（全部物质）"] + rt_analytes, key="rt_scatter_a")
            show_all = sel_rt_a == "（全部物质）"
            sel_a_single = None if show_all else sel_rt_a
            st.plotly_chart(D.single_rt_scatter(rt_df, sel_rt_date, analyte=sel_a_single),
                            width="stretch")

            # —— 保留时间波动：每物质(标准品RT + 样品RT) 的 max - min，标注最大 ——
            st.markdown("---")
            st.markdown("**保留时间波动**（每个物质：标准品+所有样品 RT 的 最大−最小）")
            qc_for_rt = qc
            sub_rt = rt_df[rt_df["exp_date"] == sel_rt_date]
            fluct = []
            for a in rt_analytes:
                sample_rts = sub_rt[sub_rt["analyte"] == a]["rt"].dropna().tolist()
                sample_rts = [r for r in sample_rts if r and r > 0]
                std_row = qc_for_rt[(qc_for_rt["exp_date"] == sel_rt_date) & (qc_for_rt["analyte"] == a)]
                std_rt = std_row["std_rt_mean"].iloc[0] if len(std_row) and pd.notna(std_row["std_rt_mean"].iloc[0]) else None
                if std_rt and std_rt > 0:
                    sample_rts.append(float(std_rt))
                if len(sample_rts) >= 1:
                    fluct.append({"物质": a, "RT波动(min)": round(max(sample_rts) - min(sample_rts), 4),
                                  "最大RT": round(max(sample_rts), 4), "最小RT": round(min(sample_rts), 4)})
            if fluct:
                fdf = pd.DataFrame(fluct)
                max_fluc = fdf["RT波动(min)"].max()
                # 柱状图，最大波动物质标红
                fig2 = go.Figure()
                fig2.add_trace(go.Bar(
                    x=fdf["物质"], y=fdf["RT波动(min)"],
                    marker_color=["#9C0006" if v == max_fluc else "#4472C4" for v in fdf["RT波动(min)"]],
                    text=[f"{v:.3f}" for v in fdf["RT波动(min)"]], textposition="outside"))
                fig2.update_layout(
                    title=f"实验 {sel_rt_date} 各物质保留时间波动（最大: {fdf.loc[fdf['RT波动(min)'].idxmax(),'物质']} = {max_fluc:.3f} min）",
                    xaxis_title="物质", yaxis_title="RT波动 (min)", template="simple_white", height=450)
                st.plotly_chart(fig2, width="stretch")
                st.caption("红色柱 = 波动最大的物质。波动 = 该物质(标准品+所有样品)保留时间的 最大值−最小值。")

    # ---- Tab5: 实验环境质控（温度/湿度/柱压 各一张折线图）----
    if qsel == "🌡️ 实验环境质控":
        st.markdown("**实验环境质控**（不同实验日期的 温度 / 湿度 / 柱压，各一张图）")
        exps_env = M.list_experiments(conn)
        for fi, fg in enumerate(D.environment_chart(exps_env)):
            st.plotly_chart(fg, width="stretch", key=f"env_fig_{fi}")

    # ---- Tab5b: 单次实验半峰宽散点（与保留时间同类：纵=半峰宽，横=样品序号）----
    if qsel == "📈 单次实验半峰宽散点":
        st.markdown("**单次实验半峰宽散点**（纵=半峰宽(min)，横=样品序号；半峰宽取自 ALLDATA 的 Width at 50% 列）")
        exps_w = M.list_experiments(conn)
        if len(exps_w) == 0:
            st.caption("暂无实验记录")
        else:
            # 独立缓存（只存 ParseResult，避免与其他页面的缓存对象互相污染）
            cache_w = st.session_state.setdefault("_width_res_cache", {})
            _map_df_w = dict(zip(exps_w["exp_date"], exps_w["folder"]))
            w_dates = sorted(exps_w["exp_date"].dropna().unique())
            sel_w_date = st.selectbox("选择实验日期", w_dates, key="width_date")
            _folder_match = _map_df_w.get(sel_w_date)
            _sub_w = pd.DataFrame()
            if _folder_match:
                _w_fmt = (_map_df_w_row_fmt := None) or None
                try:
                    _w_fmt = exps_w[exps_w["folder"] == _folder_match]["sample_format"].iloc[0]
                except Exception:
                    _w_fmt = None
                _w_key = f"{_folder_match}|{_w_fmt or ''}"
                _res_w = cache_w.get(_w_key)
                if not hasattr(_res_w, "df") or "group" not in getattr(_res_w.df, "columns", []):
                    try:
                        _res_w = P.parse_experiment(config.EXPERIMENTS_DIR / _folder_match,
                                                    sample_pattern=_w_fmt or None)
                        cache_w[_w_key] = _res_w
                    except Exception as e:
                        st.caption(f"读取失败：{e}")
                        _res_w = None
                if _res_w is not None and "group" in getattr(_res_w.df, "columns", []):
                    _sub_w = _res_w.df[(_res_w.df["category"] == "sample") &
                                       (_res_w.df["phase"] == "A") & (_res_w.df["width50"] > 0)].copy()
            if len(_sub_w) == 0:
                st.caption("（该实验无受体相半峰宽数据）")
            else:
                w_analytes = sorted(_sub_w["analyte"].unique())
                sel_w_a = st.selectbox("选择物质", ["（全部物质）"] + w_analytes, key="width_a")
                _d_w = _sub_w if sel_w_a == "（全部物质）" else _sub_w[_sub_w["analyte"] == sel_w_a]
                _d_w = _d_w.sort_values(["group", "replicate"]).reset_index(drop=True)
                _order_map = {}
                _seq = 0
                for _, rr_ in _d_w.iterrows():
                    k_ = (rr_["group"], rr_["replicate"])
                    if k_ not in _order_map:
                        _seq += 1
                        _order_map[k_] = _seq
                _d_w["样品序号"] = _d_w.apply(lambda r_: _order_map[(r_["group"], r_["replicate"])], axis=1)
                _d_w["样品"] = _d_w["sample"]
                _title_a = sel_w_a if sel_w_a != "（全部物质）" else "全部物质"
                fig_w = px.scatter(_d_w, x="样品序号", y="width50",
                                   color="analyte" if sel_w_a == "（全部物质）" else None,
                                   hover_data=["样品"],
                                   title=f"实验 {sel_w_date} 半峰宽散点 - {_title_a}",
                                   labels={"样品序号": "样品序号", "width50": "半峰宽(min)"},
                                   template="simple_white")
                fig_w.update_layout(height=460, xaxis=dict(dtick=1))
                st.plotly_chart(fig_w, width="stretch", key="width_scatter")

    # ---- 跨实验RSD超阈统计：同一分析物在多少次实验中 RSD>15% → 判断仪器/单次状态 ----
    if qsel == "📈 跨实验RSD超阈统计":
        st.markdown("**跨实验 RSD 超阈统计**（统计每個分析物在多少次实验中标准品RSD>15%）")
        st.caption("判定参考：**某一两个分析物**在多次实验中反复超阈 → 疑似该物质的**仪器响应问题**"
                   "（离子对干扰/色谱峰形）；**多个分析物只在同一两次实验**超阈 → 疑似**单次实验状态问题**"
                   "（流动相/柱温/进样重复性）。")
        _rsd_cols_stat = [("std_acceptor_rsd", "受体相标准品RSD%"), ("std_donor_rsd", "供体相标准品RSD%")]
        _rc1, _rc2 = st.columns(2)
        _sel_rsd_col = _rc1.selectbox("统计对象", ["受体相标准品", "供体相标准品", "两相合并（任一超阈计一次）"],
                                      key="qc_rsd_stat_col")
        if _sel_rsd_col == "受体相标准品":
            _stat_cols_use = [("std_acceptor_rsd", "受体相标准品RSD%")]
        elif _sel_rsd_col == "供体相标准品":
            _stat_cols_use = [("std_donor_rsd", "供体相标准品RSD%")]
        else:
            _stat_cols_use = _rsd_cols_stat
        _flag_rows = []
        for _ci, (_col, _cname) in enumerate(_stat_cols_use):
            if _col not in qc.columns:
                continue
            _sub = qc[qc[_col].notna() & (qc[_col] > config.RSD_THRESHOLD)]
            for (_a, _d), _g in _sub.groupby(["analyte", "exp_date"]):
                _flag_rows.append({"分析物": _a, "实验日期": str(_d), "对象": _cname,
                                   "RSD%": round(float(_g[_col].max()), 2)})
        if not _flag_rows:
            st.success("✅ 全部实验的所有分析物 RSD 均 ≤ 15%，无超阈记录。")
        else:
            _fdf = pd.DataFrame(_flag_rows)
            # 汇总透视：分析物 × 超阈次数
            _piv = _fdf.pivot_table(index="分析物", columns="对象", values="实验日期",
                                    aggfunc="count", fill_value=0)
            _piv["超阈总次数"] = _piv.sum(axis=1)
            _piv["涉及实验数"] = _fdf.groupby("分析物")["实验日期"].nunique()
            _piv = _piv.sort_values("超阈总次数", ascending=False)
            st.markdown("**按分析物汇总**（红色=该物质反复超阈，优先排查）")
            st.dataframe(_piv.style.map(
                lambda v: "background-color: #ffc7ce; color: #9C0006"
                if isinstance(v, (int, float)) and not pd.isna(v) and v >= 3 else ""),
                width="stretch")
            # 提示判定
            _n_multi = int((_piv["涉及实验数"] >= 3).sum())
            if _n_multi:
                st.warning(f"⚠️ 有 {_n_multi} 个分析物在 ≥3 次实验中RSD超阈（红标）——"
                           "若集中在相同物质，优先按仪器响应问题排查（离子源污染/色谱柱效/离子对干扰）。")
            else:
                st.info("超阈物质均未跨 ≥3 次实验重复出现——更可能是单次实验状态问题，"
                        "结合上方明细查看对应日期的其他物质是否同时超阈。")
            with st.expander("查看超阈明细（逐条）"):
                st.dataframe(_fdf.sort_values(["分析物", "实验日期"]),
                             width="stretch", hide_index=True)
            # 每日期超阈物质数（单次状态问题的信号）
            _per_date = _fdf.groupby("实验日期")["分析物"].nunique().sort_values(ascending=False)
            if len(_per_date):
                _wd = _per_date.index[0]
                st.caption(f"超阈分析物最多的实验日期：{_wd}（{_per_date.iloc[0]} 个物质）"
                           + ("——该日期多个物质同时超阈，符合单次状态问题特征。" if _per_date.iloc[0] >= 5 else ""))

    # ---- Tab6: 异常统计（渗漏/烧焦/破裂；按SLM考察试剂稳定 + 96孔板定位排查）----
    if qsel == "⚠️ 异常统计(渗漏/烧焦/破裂)":
        exps_env = M.list_experiments(conn)  # 自取数据（不依赖其他分区是否渲染）
        st.markdown("**异常统计**（汇总所有实验的异常标签，按 SLM 统计渗漏/烧焦/破裂次数考察试剂稳定性，"
                    "并结合 96 孔板位置排查是试剂问题还是电极棒问题）")
        # 收集所有实验的异常记录
        all_abn = []
        for _, erow in exps_env.iterrows():
            eid_a = int(erow["id"])
            gcond_a = M.get_group_conditions(conn, eid_a)
            for _, gr in gcond_a.iterrows():
                if gr["grp"] == 0:
                    continue
                try:
                    tags_a = json.loads(gr["note_tags"]) if gr["note_tags"] else []
                except Exception:
                    tags_a = []
                if tags_a:
                    # 孔位：异常孔位优先（精细记录）；否则该组使用的全部孔位
                    wells_used = M.parse_wells(gr["well_position"])
                    wells_abn = M.parse_wells(gr["abnormal_wells"])
                    wells_show = wells_abn if wells_abn else wells_used
                    all_abn.append({
                        "实验日期": erow["exp_date"],
                        "组别": gr["group_label"] or f"M{int(gr['grp'])}",
                        "主SLM": gr["slm_main"] or "（未记录）",
                        "孔位": ", ".join(wells_show),
                        "标签": tags_a,
                    })
        if not all_abn:
            st.success("✅ 目前所有实验都没有异常记录（录入实验时在「异常标签」里勾选渗漏/烧焦/破裂即可统计）。")
        else:
            abn_df = pd.DataFrame(all_abn)
            # 展开 tags 到行
            exp_rows = []
            for _, r in abn_df.iterrows():
                for t in r["标签"]:
                    exp_rows.append({"实验日期": r["实验日期"], "组别": r["组别"], "主SLM": r["主SLM"],
                                     "孔位": r["孔位"], "异常类型": t})
            flat = pd.DataFrame(exp_rows)
            # 按SLM统计
            st.markdown("**按 SLM 统计**（同一 SLM 异常次数多 → 提示试剂稳定性问题）")
            piv = flat.pivot_table(index="主SLM", columns="异常类型", values="组别",
                                   aggfunc="count", fill_value=0)
            piv["异常总次数"] = piv.sum(axis=1)
            # 涉及实验次数
            n_exp = flat.groupby("主SLM")["实验日期"].nunique()
            piv.insert(0, "涉及实验数", n_exp)
            st.dataframe(piv, width="stretch")
            # 明细
            with st.expander("查看异常明细（逐条）"):
                st.dataframe(flat.sort_values(["主SLM", "实验日期"]),
                             width="stretch", hide_index=True)
            # 96孔板视图
            st.markdown("---")
            st.markdown("**96 孔板位置视图**（异常按孔位分布——同一列/行反复异常 → 疑似电极棒问题；"
                        "同一试剂多个位置异常 → 疑似试剂问题）")
            tagged = flat[flat["孔位"] != ""]
            if len(tagged) == 0:
                st.caption("（异常记录中未填写孔位——录入实验时在每组选「孔位」即可在此定位）")
            else:
                TAG_COLOR = {"渗漏": "#E45756", "烧焦": "#F28E2B", "破裂": "#EDC948"}
                wells = [f"{r}{c}" for r in "ABCDEFGH" for c in range(1, 13)]
                well_map = {w: [] for w in wells}
                for _, r in tagged.iterrows():
                    # 孔位字段为逗号分隔多孔，逐孔展开
                    for w in [x.strip() for x in str(r["孔位"]).split(",") if x.strip()]:
                        if w in well_map:
                            well_map[w].append(r)
                fig_well = go.Figure()
                xs, ys, colors, texts = [], [], [], []
                for w, recs in well_map.items():
                    if not recs:
                        continue
                    row_l, col_n = w[0], int(w[1:])
                    xs.append(col_n)
                    ys.append(row_l)
                    first_tag = recs[0]["异常类型"]
                    colors.append(TAG_COLOR.get(first_tag, "#999999"))
                    texts.append("<br>".join(f"{x['实验日期']} {x['组别']}: {x['异常类型']}" for x in recs))
                fig_well.add_trace(go.Scatter(
                    x=xs, y=ys, mode="markers+text",
                    marker=dict(size=38, color=colors, line=dict(color="#333", width=1)),
                    text=[w for w, recs in well_map.items() if recs],
                    textfont=dict(color="white", size=10),
                    hovertemplate="%{text}<extra></extra>"))
                for t, c in TAG_COLOR.items():
                    fig_well.add_trace(go.Scatter(
                        x=[None], y=[None], mode="markers",
                        marker=dict(size=12, color=c), name=t, legendgroup=t))
                fig_well.update_layout(
                    title="96孔板异常位置分布（颜色=异常类型；悬浮显示日期/组别/类型）",
                    xaxis=dict(title="列 (1-12)", dtick=1, side="top"),
                    yaxis=dict(title="行 (A-H)", autorange="reversed"),
                    template="simple_white", height=420)
                st.plotly_chart(fig_well, width="stretch", key="well_fig")


# ========================================================
# 4. 自定义看板（物质性质 vs 指标；多面板对比）
# ========================================================
def page_custom_dashboard():
    st.title("🔧 自定义看板")
    conn = get_db()
    mdf = M.get_metrics_df(conn)
    exps = M.list_experiments(conn)
    if len(mdf) == 0:
        st.info("暂无数据。"); return

    analyte_cols = get_analyte_numeric_cols() if REF_OK else []
    # 合并组级条件 + 分析物性质
    gc_all = pd.read_sql(
        "SELECT g.experiment_id, g.grp, g.group_label, g.slm_main, g.pH_donor, g.pH_acceptor, "
        "g.extract_time, g.extract_voltage FROM group_conditions g WHERE g.grp>0", conn)
    mdf_ext = mdf.merge(gc_all, left_on=["experiment_id", "grp"], right_on=["experiment_id", "grp"],
                        how="left", suffixes=("", "_g"))
    if REF_OK and analyte_cols:
        ab = R.get_analyte_basic()
        if ab is not None:
            name_col = "Englishname" if "Englishname" in ab.columns else ab.columns[0]
            ab_keyed = ab[[name_col] + analyte_cols].copy()
            ab_keyed[name_col] = ab_keyed[name_col].map(lambda x: R._norm_analyte_name(str(x)))
            ab_keyed = ab_keyed.set_index(name_col)
            mdf_ext["_norm"] = mdf_ext["analyte"].map(R._norm_analyte_name)
            if not ab_keyed.index.is_unique:
                ab_keyed = ab_keyed[~ab_keyed.index.duplicated(keep="first")]  # 主键重名时防 join 行倍增
            mdf_ext = mdf_ext.join(ab_keyed, on="_norm", how="left")

    # 按供体相pH的属性列（charge/LogD/mg/ml 等，每行按该组pH现查）→ 作为 X/Y 候选
    ph_derived_cols = {}
    if REF_OK:
        # 唯一 (物质, pH) 对各查一次（get_property_by_ph 本身有 lru，这里再省掉逐行 apply 开销）
        _uniq_pairs = mdf_ext[["analyte", "pH_donor"]].drop_duplicates()
        for key, meta_p in R.list_ph_properties().items():
            disp = meta_p["display"]
            col_name = f"{disp}@供体pH"
            _pair_val = {}
            for _a, _ph in zip(_uniq_pairs["analyte"], _uniq_pairs["pH_donor"]):
                try:
                    _phf = float(_ph)
                    if _phf != _phf:  # NaN：float() 不抛但派生查询会取到任意 pH 的值
                        raise ValueError("NaN")
                except (TypeError, ValueError):
                    _pair_val[(_a, _ph)] = None
                    continue
                _pair_val[(_a, _ph)] = R.get_property_by_ph(_a, disp, _phf)
            mdf_ext[col_name] = [_pair_val[(_a, _ph)]
                                 for _a, _ph in zip(mdf_ext["analyte"], mdf_ext["pH_donor"])]
            ph_derived_cols[col_name] = col_name

    # 构造"实验日期+组"选项
    exp_group_options = []
    for _, r in mdf_ext[["experiment_id", "exp_date", "grp", "group_label", "slm_main"]].drop_duplicates().iterrows():
        lbl = r["group_label"] or f"M{r['grp']}"
        exp_group_options.append({
            "key": f"{r['exp_date']}|M{r['grp']}",
            "label": f"{r['exp_date']} - {lbl}",
            "exp_date": r["exp_date"], "grp": r["grp"], "slm_main": r["slm_main"],
        })

    # 偏好持久化
    st.markdown("##### 偏好保存")
    prefs = M.list_prefs(conn)
    pc1, pc2, pc3 = st.columns([2, 1, 1])
    pref_name = pc1.text_input("偏好名", value="我的看板", key="pref_name")
    if pc2.button("💾 保存当前配置"):
        M.save_pref(conn, pref_name, st.session_state.get("dash_cfg", []))
        st.success(f"已保存偏好「{pref_name}」")
    if pc3.button("📂 加载") and pref_name in prefs:
        st.session_state["dash_cfg"] = M.load_pref(conn, pref_name)
        st.rerun()

    # ============ 面板模型：每个面板 = 一个实验组 + 多组XY（+号添加）============
    # X 轴候选：分析物性质列 + 按供体pH派生列（charge/LogD/mg-ml 等）
    x_candidates = analyte_cols + list(ph_derived_cols.keys())
    x_candidates = x_candidates or ["Logp"]
    # Y 轴候选：内置指标
    y_label_options = [METRIC_LABELS[k] for k in ["recovery", "dm", "ma", "A_rsd", "S_rsd"]]
    default_xy = {"x": analyte_cols[0] if analyte_cols else "Logp", "y": "recovery"}
    default_grp = exp_group_options[0]["key"] if exp_group_options else ""

    import uuid as _uuid

    def _dash_nid():
        """面板/XY 的稳定短 id：widget 键用 id 而非索引，加载偏好/删除面板时不再错位"""
        return _uuid.uuid4().hex[:8]

    if "dash_cfg" not in st.session_state:
        st.session_state["dash_cfg"] = [{"_id": _dash_nid(), "exp_grp": default_grp,
                                         "xys": [{**dict(default_xy), "_id": _dash_nid()}]}]
    # 兼容旧配置（无 xys 字段 / 无 _id）
    for p in st.session_state["dash_cfg"]:
        if "xys" not in p:
            p["xys"] = [{"x": p.get("x", default_xy["x"]), "y": p.get("y", "recovery")}]
        p.setdefault("_id", _dash_nid())
        for _xy in p["xys"]:
            _xy.setdefault("_id", _dash_nid())
    cfg = st.session_state["dash_cfg"]

    cc = st.columns([1, 1, 2])
    if cc[0].button("➕ 添加面板（新实验组）"):
        cfg.append({"_id": _dash_nid(), "exp_grp": default_grp,
                    "xys": [{**dict(default_xy), "_id": _dash_nid()}]})
        st.rerun()
    if cc[1].button("➖ 删除最后面板") and len(cfg) > 1:
        cfg.pop(); st.rerun()
    st.caption(f"当前 {len(cfg)} 个面板，每个面板绑定一个实验组，组内可添加多组XY")

    opts = [o["key"] for o in exp_group_options]
    def _grp_label(k):
        return next((o["label"] for o in exp_group_options if o["key"] == k), k)

    # 配置每个面板
    for i, panel in enumerate(cfg):
        cur = panel.get("exp_grp", opts[0] if opts else "")
        panel["exp_grp"] = cur  # 防空
        with st.expander(f"面板 {i+1}：{_grp_label(panel['exp_grp'])}", expanded=(i == 0)):
            pcc0 = st.columns([3, 1])
            panel["exp_grp"] = pcc0[0].selectbox("实验组", opts,
                index=opts.index(cur) if cur in opts else 0, key=f"p{panel['_id']}_eg",
                format_func=_grp_label)
            if pcc0[1].button("🗑 删除此面板", key=f"p{panel['_id']}_del") and len(cfg) > 1:
                cfg.pop(i); st.rerun()
            # 组内多组 XY
            # Y 轴映射：内置指标(中文标签→列名) + 自定义列(峰高均值→自身)
            y_map = dict(LABEL_TO_METRIC)
            for _c in y_label_options:
                if _c not in y_map:
                    y_map[_c] = _c
            xys = panel.setdefault("xys", [dict(default_xy)])
            for j, xy in enumerate(xys):
                xyc = st.columns([2, 2, 1])
                xy["x"] = xyc[0].selectbox(f"X轴 {j+1}", x_candidates,
                    index=x_candidates.index(xy.get("x")) if xy.get("x") in x_candidates else 0,
                    key=f"p{panel['_id']}_xy{xy['_id']}_x")
                cur_y_lbl = xy.get("y", "recovery")
                if cur_y_lbl not in y_label_options:
                    cur_y_lbl = METRIC_LABELS.get(cur_y_lbl, y_label_options[0])
                cur_y_lbl = xyc[1].selectbox(f"Y轴 {j+1}", y_label_options,
                    index=y_label_options.index(cur_y_lbl) if cur_y_lbl in y_label_options else 0,
                    key=f"p{panel['_id']}_xy{xy['_id']}_y")
                xy["y"] = y_map[cur_y_lbl]
                if xyc[2].button("➖", key=f"p{panel['_id']}_xy{xy['_id']}_rm", help="删除这组XY") and len(xys) > 1:
                    xys.pop(j); st.rerun()
            if st.button("➕ 添加XY", key=f"p{panel['_id']}_addxy"):
                xys.append({**dict(default_xy), "_id": _dash_nid()}); st.rerun()

    # 渲染面板
    st.markdown("---")
    ua1, ua2, ua3 = st.columns(3)
    unify_axis = ua1.checkbox("统一 Y 轴范围（所有面板）", value=False, key="unify_y")
    uy_auto = ua2.checkbox("Y轴自动范围", value=not unify_axis, key="uy_auto", disabled=not unify_axis)
    if unify_axis and not uy_auto:
        ur1, ur2 = st.columns(2)
        uy_min = ur1.number_input("Y轴最小值", value=0.0, key="uy_min")
        uy_max = ur2.number_input("Y轴最大值", value=100.0, key="uy_max")
    else:
        uy_min, uy_max = None, None

    # 组→颜色：按面板中组首次出现顺序依次分配色板（不同组必不同色，Set1色差大）
    GROUP_COLORS = (px.colors.qualitative.Set1 + px.colors.qualitative.Dark2 +
                    px.colors.qualitative.Set2)
    grp_order_local = []
    for p in cfg:
        k = p.get("exp_grp")
        if k and k not in grp_order_local:
            grp_order_local.append(k)
    _grp_color_map = {k: GROUP_COLORS[idx % len(GROUP_COLORS)] for idx, k in enumerate(grp_order_local)}
    def _grp_color(k) -> str:
        return _grp_color_map.get(k, "#4472C4")

    fig_counter = 0
    for i, panel in enumerate(cfg):
        eg = next((o for o in exp_group_options if o["key"] == panel["exp_grp"]), None)
        if eg is None:
            continue
        d_panel = mdf_ext[(mdf_ext["exp_date"] == eg["exp_date"]) & (mdf_ext["grp"] == eg["grp"])].copy()
        color_this = _grp_color(panel["exp_grp"])
        with st.container():
            st.markdown(f"**面板 {i+1}：{eg['label']}**")
            for j, xy in enumerate(panel.get("xys", [])):
                xcol, ycol = xy.get("x"), xy.get("y")
                d = d_panel.copy()
                d["_x"] = d[xcol].map(R._try_float if REF_OK else float) if xcol in d.columns else None
                fig = go.Figure()
                fig.add_trace(go.Scatter(
                    x=d["_x"], y=d[ycol], mode="markers",
                    marker=dict(color=color_this, size=9, line=dict(color="white", width=1)),
                    text=d["analyte"], hovertemplate="%{text}<extra></extra>",
                    name=eg["label"]))
                if unify_axis and uy_min is not None and uy_max is not None:
                    fig.update_yaxes(range=[uy_min, uy_max])
                fig.update_layout(
                    title=f"{METRIC_LABELS.get(ycol, ycol)} vs {xcol}",
                    xaxis_title=xcol, yaxis_title=METRIC_LABELS.get(ycol, ycol),
                    template="simple_white", height=340,
                    margin=dict(l=40, r=20, t=45, b=40),
                    showlegend=True, legend=dict(orientation="h", y=1.02, x=0))
                st.plotly_chart(fig, width="stretch", key=f"dash_fig_{fig_counter}")
                fig_counter += 1

    # ============ 详情下拉框放本页侧边栏（避免页面长难滚动）============
    analytes_all = sorted(mdf_ext["analyte"].unique())
    # 收集所有实验组的主SLM + SLM添加剂/供体相添加剂/受体相添加剂
    all_reagents = set(o["slm_main"] for o in exp_group_options if o["slm_main"])
    try:
        gc_all_add = pd.read_sql(
            "SELECT slm_additives, donor_additives, acceptor_additives FROM group_conditions", conn)
        for c in ["slm_additives", "donor_additives", "acceptor_additives"]:
            for v in gc_all_add[c].dropna():
                try:
                    items = json.loads(v)
                    all_reagents.update(x for x in items if x)
                except Exception:
                    pass
    except Exception:
        pass
    all_reagents = sorted(all_reagents)
    # 取物质所在实验组的供体相pH（用于物质详情侧栏按pH显示LogD/电荷）
    def _ph_of_analyte(a):
        sub = mdf_ext[mdf_ext["analyte"] == a]
        if "pH_donor" in sub.columns:
            phs = pd.to_numeric(sub["pH_donor"], errors="coerce").dropna()
            if len(phs):
                return float(phs.iloc[0])
        return None
    render_detail_sidebar(ph_lookup=lambda a: _ph_of_analyte(a))


# ========================================================
# 5. 实验汇总（两层表头 + 分3表）
# ========================================================
@st.cache_data(show_spinner=False)
def _build_summary_data(mdf: pd.DataFrame, exps_min: pd.DataFrame,
                        gcond_all: pd.DataFrame, excluded_map_json: str):
    """构建实验汇总数据（重型循环只在此执行；筛选交互命中缓存不重算）。
    返回 (sdf 长表, exp_rep_info {日期:(重复数,{组别名:剔除数})}, all_reagents 全部试剂名)。
    """
    labels_by_eid, cond_by_eid = {}, {}
    for _, gr in gcond_all.iterrows():
        eid = int(gr["experiment_id"])
        if int(gr["grp"]) == 0:
            continue
        adds = []
        for c in ["slm_additives", "donor_additives", "acceptor_additives"]:
            try:
                adds += [x for x in json.loads(gr[c]) if x] if gr[c] else []
            except Exception:
                pass
        labels_by_eid.setdefault(eid, {})[int(gr["grp"])] = gr.get("group_label") or f"M{int(gr['grp'])}"
        cond_by_eid.setdefault(eid, {})[int(gr["grp"])] = (gr.get("slm_main") or "", adds)
    try:
        excluded_map = json.loads(excluded_map_json)
    except Exception:
        excluded_map = {}

    from collections import Counter as _Counter
    _date_cnt = _Counter(exps_min["exp_date"].tolist())
    records = []
    for _, erow in exps_min.iterrows():
        eid = int(erow["id"])
        # 同一天多次实验：日期标签追加文件夹名，避免透视列合并与 exp_rep_info 相互覆盖
        _dlbl = erow["exp_date"] + (f"（{erow['folder']}）" if _date_cnt.get(erow["exp_date"], 0) > 1 else "")
        labels = labels_by_eid.get(eid, {})
        cond_by_grp = cond_by_eid.get(eid, {})
        sub = mdf[mdf["experiment_id"] == eid]
        for grp, gdf in sub.groupby("grp"):
            label = labels.get(int(grp), f"M{grp}")
            slm_main_g, adds_g = cond_by_grp.get(int(grp), ("", []))
            for analyte, adf in gdf.groupby("analyte"):
                records.append({
                    "实验日期": _dlbl,
                    "组别": label, "grp": int(grp), "物质": analyte,
                    "主SLM": slm_main_g,
                    "添加剂": ", ".join(adds_g) if adds_g else "（无）",
                    "recovery": adf["recovery"].iloc[0] if pd.notna(adf["recovery"].iloc[0]) else None,
                    "dm": adf["dm"].iloc[0] if pd.notna(adf["dm"].iloc[0]) else None,
                    "ma": adf["ma"].iloc[0] if pd.notna(adf["ma"].iloc[0]) else None,
                    "A_rsd": adf["A_rsd"].iloc[0] if pd.notna(adf["A_rsd"].iloc[0]) else None,
                    "S_rsd": adf["S_rsd"].iloc[0] if pd.notna(adf["S_rsd"].iloc[0]) else None,
                })
    sdf = pd.DataFrame(records)

    exp_rep_info = {}
    for _, erow in exps_min.iterrows():
        eid = int(erow["id"])
        n_rep_s = int(erow["n_replicates"]) if pd.notna(erow["n_replicates"]) else 3
        label_by_num = labels_by_eid.get(eid, {})
        excl_by_label = {}
        _fmt_s = erow.get("sample_format") if hasattr(erow, "get") else None
        _fmt_s = _fmt_s if isinstance(_fmt_s, str) and _fmt_s else None
        for s_ in excluded_map.get(str(eid), []):
            _pp_s = P.parse_sample_name(str(s_), pattern=_fmt_s)
            if _pp_s:
                lbl = label_by_num.get(int(_pp_s[0]))
                if lbl:
                    excl_by_label[lbl] = excl_by_label.get(lbl, 0) + 1
        _dlbl2 = erow["exp_date"] + (f"（{erow['folder']}）" if _date_cnt.get(erow["exp_date"], 0) > 1 else "")
        exp_rep_info[_dlbl2] = (n_rep_s, excl_by_label)

    all_reagents = sorted({v for d in cond_by_eid.values() for (slm, adds) in d.values()
                           for v in ([slm] if slm else []) + adds})
    return sdf, exp_rep_info, all_reagents


def page_summary():
    st.title("📈 实验汇总")
    conn = get_db()
    exps = M.list_experiments(conn)
    mdf = M.get_metrics_df(conn)
    if len(exps) == 0:
        st.info("暂无实验记录。请到「新增实验」录入。"); return

    # —— 实验列表（总览表）——
    st.subheader("📋 实验列表")
    show_cols = ["exp_date", "folder", "exp_type", "n_groups", "temperature", "humidity", "purpose"]
    show_cols = [c for c in show_cols if c in exps.columns]
    st.dataframe(exps[show_cols], width="stretch", hide_index=True)
    st.markdown("---")

    if len(mdf) == 0:
        st.info("暂无实验数据。"); return

    # 汇总数据构建（缓存：仅当实验/指标数据变化时重算，筛选交互直接命中缓存不再卡顿）
    gcond_all = pd.read_sql(
        "SELECT experiment_id, grp, group_label, slm_main, slm_additives, pH_donor, pH_acceptor, "
        "donor_additives, acceptor_additives FROM group_conditions WHERE grp>0", conn)
    excluded_map = {str(eid_): sorted(excl_set)
                    for eid_, excl_set in M.get_all_excluded_samples(conn).items()}
    exps_min = exps[["id", "exp_date", "n_replicates", "folder", "sample_format"]]
    sdf, exp_rep_info, all_reagents_sum = _build_summary_data(mdf, exps_min, gcond_all, json.dumps(excluded_map))
    if len(sdf) == 0:
        st.info("暂无实验数据（实验汇总需要先录入实验并生成数据）。")
        return

    # 排序：按分析物性质
    sort_cols = get_analyte_numeric_cols() if REF_OK else []
    sc1, sc2 = st.columns(2)
    sort_by = sc1.selectbox("按分析物性质排序（物质行顺序）", ["（不排序）"] + sort_cols, key="sum_sort")
    sort_asc = sc2.checkbox("升序", value=True, key="sum_asc")
    analyte_order = sorted(sdf["物质"].unique())
    if sort_by != "（不排序）" and REF_OK:
        ab = R.get_analyte_basic()
        if ab is not None:
            name_col = "Englishname" if "Englishname" in ab.columns else ab.columns[0]
            prop_map = dict(zip(ab[name_col].map(lambda x: R._norm_analyte_name(str(x))),
                                ab[sort_by].map(R._try_float)))
            analyte_order = sorted(analyte_order,
                key=lambda a: (prop_map.get(R._norm_analyte_name(a)) if prop_map.get(R._norm_analyte_name(a)) is not None else 1e9))
            if not sort_asc:
                analyte_order = list(reversed(analyte_order))

    # 筛选 + 指标表（fragment 局部刷新：筛选/配色/指标切换只重跑本区块，不再整页重渲染）
    _summary_tables_fragment(sdf, analyte_order, exp_rep_info)

    # 详情侧栏（三页共用，选择跨页保留）
    def _ph_lookup_sum(a):
        sub = sdf[sdf["物质"] == a]
        if len(sub) == 0:
            return None
        exp_date = sub["实验日期"].iloc[0]
        erow = exps[exps["exp_date"] == exp_date]
        if len(erow):
            eid_s = int(erow.iloc[0]["id"])
            phs = pd.to_numeric(
                gcond_all[gcond_all["experiment_id"] == eid_s]["pH_donor"],
                errors="coerce").dropna()
            return float(phs.iloc[0]) if len(phs) else None
        return None
    render_detail_sidebar(ph_lookup=_ph_lookup_sum)


@st.fragment
def _summary_tables_fragment(sdf: pd.DataFrame, analyte_order: list, exp_rep_info: dict):
    """实验汇总的筛选+表格区（fragment 局部刷新，解决筛选卡顿）；含配色偏好记忆。"""
    # 筛选：日期 / 主SLM / 添加剂 / 分析物（组别筛选改为按条件筛选）
    fc1, fc2, fc3, fc4 = st.columns(4)
    sel_dates = fc1.multiselect("筛选实验日期", sorted(sdf["实验日期"].unique()), key="sum_fd")
    all_slms_sum = sorted(set(x for x in sdf["主SLM"] if x))
    sel_slms_f = fc2.multiselect("按主SLM筛选", all_slms_sum, key="sum_fsm")
    all_adds_sum = sorted(set(a for s in sdf["添加剂"] for a in s.split(", ") if a != "（无）"))
    sel_adds_f = fc3.multiselect("按添加剂筛选", all_adds_sum, key="sum_fadd")
    sel_analytes = fc4.multiselect("筛选分析物", sorted(sdf["物质"].unique()), key="sum_fa")
    view = sdf.copy()
    if sel_dates: view = view[view["实验日期"].isin(sel_dates)]
    if sel_slms_f: view = view[view["主SLM"].isin(sel_slms_f)]
    if sel_adds_f:
        view = view[view["添加剂"].apply(lambda s: any(a in s.split(", ") for a in sel_adds_f))]
    if sel_analytes:
        view = view[view["物质"].isin(sel_analytes)]
        analyte_order = [a for a in analyte_order if a in sel_analytes]

    # RSD 显示参数
    show_rsd = st.checkbox("显示 RSD（格式：值 (RSD:xx%)）", value=True, key="sum_rsd")
    # 颜色映射参数（配色偏好自动记忆：选择即保存，下次打开默认为上次用的配色）
    cm1, cm2, cm3 = st.columns(3)
    color_on = cm1.checkbox("启用颜色映射", value=True, key="sum_color")
    _cmap_pref = S.get_cmap_pref()
    _cmap_choices = ["RdYlGn", "viridis", "plasma", "Blues", "RdBu", "YlGnBu"]
    cmap_sum = cm2.selectbox(
        "配色（自动记住上次选择）", _cmap_choices,
        index=_cmap_choices.index(_cmap_pref) if _cmap_pref in _cmap_choices else 0,
        key="sum_cmap",
        on_change=lambda: S.save_cmap_pref(st.session_state.get("sum_cmap", "RdYlGn"))) if color_on else "RdYlGn"
    cv_range = cm3.checkbox("自定义颜色范围", value=False, key="sum_crange")
    if cv_range:
        cr1, cr2 = st.columns(2)
        cmin = cr1.number_input("颜色最小值", value=0.0, key="sum_cmin")
        cmax = cr2.number_input("颜色最大值", value=100.0, key="sum_cmax")
    else:
        cmin, cmax = None, None

    # 指标切换（单选代替三标签页：只渲染当前查看的表，速度更快）
    metric_tabs = [("recovery", "回收率%"), ("dm", "供体相至SLM转移率%"), ("ma", "SLM至受体相转移率%")]
    _labels = [t[1] for t in metric_tabs]
    msel = st.radio("查看指标", _labels, horizontal=True, key="sum_metric_sel")
    mcol, mtitle = metric_tabs[_labels.index(msel)]
    # 展开全部药物行：表格高度=完整内容高度（页面滚动看行），只有组列过多时才在表内横向滑动
    expand_all = st.checkbox("展开全部药物行（表格完整高度，行不再在表格内滚动）",
                             value=True, key="sum_expand_all")
    _render_two_level_table(view, mcol, mtitle, analyte_order, show_rsd, color_on, cmap_sum, cmin, cmax,
                            exp_rep_info=exp_rep_info, full_height=expand_all)


def _render_two_level_table(df: pd.DataFrame, metric_col: str, title: str, analyte_order: list,
                            show_rsd: bool = True, color_on: bool = True,
                            cmap_name: str = "RdYlGn", vmin=None, vmax=None,
                            exp_rep_info: dict = None, full_height: bool = False):
    """exp_rep_info: {实验日期: (n_replicates, excluded_set)} 用于回收率表的有效膜数行"""
    """两层表头：第一层实验日期，第二层组别名；行=物质。
    颜色映射与 RSD 显示可并行：
    - color_on：背景按数值 colormap 着色（优先）；show_rsd 时文本附 (RSD:xx%)
    - 仅 show_rsd（color_on=False）时：RSD>15% 标红
    """
    if len(df) == 0:
        st.caption("无数据"); return
    pivot = df.pivot_table(index="物质", columns=["实验日期", "组别"], values=metric_col, aggfunc="first")
    pivot = pivot.reindex([a for a in analyte_order if a in pivot.index])
    n_analytes_sum = len(pivot)  # 数据行数（统计行追加前记录）
    rsd_col = "S_rsd" if metric_col in ("dm", "ma") else "A_rsd"
    pivot_rsd = df.pivot_table(index="物质", columns=["实验日期", "组别"], values=rsd_col, aggfunc="first")
    pivot_rsd = pivot_rsd.reindex(pivot.index)

    # 色标范围
    vals = pivot.values.flatten()
    vals = vals[~pd.isna(vals)]
    if vmin is None:
        vmin = float(np.min(vals)) if len(vals) else 0.0
    if vmax is None:
        vmax = float(np.max(vals)) if len(vals) else 100.0
    if vmin == vmax:
        vmax = vmin + 1.0

    # 背景颜色（色标设置；显示文本与样式在统计行生成后统一向量化构建）
    import matplotlib as mpl
    cmap_obj = None
    if color_on:
        try:
            cmap_obj = mpl.colormaps[cmap_name]
        except Exception:
            cmap_obj = mpl.colormaps["RdYlGn"]
        norm = mpl.colors.Normalize(vmin=vmin, vmax=vmax)

    # —— 按实验组统计行（回收率/转移率适用；RSD 参考线固定 >15%）——
    # 参考线默认值按指标分开（回收率/转移率各自独立设置与保存）
    ref_kind_sum = "recovery" if metric_col == "recovery" else "transfer"
    _s_r1, _s_r2 = S.get_ref_lines(ref_kind_sum)
    rate_cols = {"recovery": None, "dm": None, "ma": None}
    if metric_col in rate_cols:
        # 参考线设置（与实验详情热图一致：用户可设1~2条，可保存为该指标默认）
        with st.expander(f"📐 {title} 统计参考线设置（决定底部按实验组统计的分档）"):
            use_ref = st.checkbox("启用统计参考线", value=(_s_r1 is not None),
                                  key=f"sum_ref_on_{metric_col}")
            ref1 = ref2 = None
            if use_ref:
                rr1, rr2 = st.columns(2)
                ref1 = rr1.number_input("参考线1",
                                        value=float(_s_r1 if _s_r1 is not None else 20.0),
                                        key=f"sum_ref1_{metric_col}")
                use_ref2 = rr2.checkbox("启用第2条参考线", value=_s_r2 is not None,
                                        key=f"sum_ref2_on_{metric_col}")
                if use_ref2:
                    ref2 = rr2.number_input("参考线2",
                                            value=float(_s_r2 if _s_r2 is not None else 40.0),
                                            key=f"sum_ref2_{metric_col}")
                if st.button("💾 保存为该指标默认参考线", key=f"sum_ref_save_{metric_col}"):
                    S.save_ref_lines(ref1 if ref1 is not None else _s_r1,
                                     ref2 if ref2 is not None else _s_r2,
                                     kind=ref_kind_sum)
                    st.success("已保存为该指标默认参考线。")
        if use_ref and ref1 is not None:
            if ref2 is not None and ref2 > ref1:
                lo, hi = ref1, ref2
                row_lo = [int((pivot[c] < lo).sum()) for c in pivot.columns]
                row_mid = [int(((pivot[c] >= lo) & (pivot[c] < hi)).sum()) for c in pivot.columns]
                row_hi = [int((pivot[c] >= hi).sum()) for c in pivot.columns]
                pivot.loc[f"📊 <{lo:g} 个数"] = row_lo
                pivot.loc[f"📊 {lo:g}~{hi:g} 个数"] = row_mid
                pivot.loc[f"📊 ≥{hi:g} 个数"] = row_hi
                stat_desc = f"底部按实验组统计：<{lo:g} / {lo:g}~{hi:g} / ≥{hi:g} 的个数"
            else:
                row_lo = [int((pivot[c] < ref1).sum()) for c in pivot.columns]
                row_hi = [int((pivot[c] >= ref1).sum()) for c in pivot.columns]
                pivot.loc[f"📊 <{ref1:g} 个数"] = row_lo
                pivot.loc[f"📊 ≥{ref1:g} 个数"] = row_hi
                stat_desc = f"底部按实验组统计：<{ref1:g} / ≥{ref1:g} 的个数"
        else:
            if metric_col == "recovery":
                # 有效膜数 = 重复数 − 该组被剔除样本数（按列的 实验日期+组别名 取对应剔除数）
                exp_rep_info = exp_rep_info or {}
                row_valid = []
                for c in pivot.columns:
                    exp_d = c[0] if isinstance(c, tuple) else c
                    g_lbl = c[1] if isinstance(c, tuple) else c
                    n_rep, excl_by_label = exp_rep_info.get(exp_d, (3, {}))
                    row_valid.append(max(0, int(n_rep) - excl_by_label.get(g_lbl, 0)))
                pivot.loc["📊 有效膜数"] = row_valid
                stat_desc = "未启用参考线。底部行为各实验组有效膜数（重复数−被剔除数）"
            else:
                stat_desc = "未启用参考线（可在上方参考线设置中启用以显示分档统计）"
        # 回收率：无论参考线是否启用，底部都显示有效膜数行
        if metric_col == "recovery" and "📊 有效膜数" not in pivot.index:
            exp_rep_info = exp_rep_info or {}
            row_valid = []
            for c in pivot.columns:
                exp_d = c[0] if isinstance(c, tuple) else c
                g_lbl = c[1] if isinstance(c, tuple) else c
                n_rep, excl_by_label = exp_rep_info.get(exp_d, (3, {}))
                row_valid.append(max(0, int(n_rep) - excl_by_label.get(g_lbl, 0)))
            pivot.loc["📊 有效膜数"] = row_valid
            stat_desc = (stat_desc + "；底部末行为各实验组有效膜数（重复数−被剔除数）"
                         if "stat_desc" in dir() else "底部末行为各实验组有效膜数（重复数−被剔除数）")
        # 显示文本 + 单元格样式一次向量化生成（此前逐单元格 .loc 在 32物质×多组列时
        # 是数万次 pandas 标量操作，是筛选卡顿的另一半来源）；统计行白底直接写入样式矩阵
        _pv = pivot.to_numpy(dtype=float)
        _rv = pivot_rsd.reindex(index=pivot.index, columns=pivot.columns).to_numpy(dtype=float)
        _nr, _nc = _pv.shape
        _disp = np.empty((_nr, _nc), dtype=object)
        _css = np.empty((_nr, _nc), dtype=object)
        for _i in range(_nr):
            if str(pivot.index[_i]).startswith("📊"):
                for _j in range(_nc):
                    _v_stat = pivot.iloc[_i, _j]
                    # 统计行整数显示（11 而非 11.0）
                    _disp[_i, _j] = str(int(_v_stat)) if float(_v_stat).is_integer() else str(_v_stat)
                    _css[_i, _j] = "background-color: #ffffff; color: #000000"
                continue
            for _j in range(_nc):
                _v = _pv[_i, _j]
                _r = _rv[_i, _j]
                if np.isnan(_v):
                    _disp[_i, _j] = "-"
                    _css[_i, _j] = ""
                else:
                    _s = f"{_v:.2f}"
                    if show_rsd and not np.isnan(_r):
                        _s += f" (RSD:{_r:.1f}%)"
                    _disp[_i, _j] = _s
                    if color_on:
                        # 颜色映射优先（RSD 不标红）
                        _css[_i, _j] = f"background-color: {mpl.colors.to_hex(cmap_obj(norm(_v)))}"
                    else:
                        # 无颜色映射时才用 RSD>15% 标红
                        _css[_i, _j] = ("background-color: #ffc7ce; color: #9C0006"
                                        if (not np.isnan(_r) and _r > 15.0) else "")
        display = pd.DataFrame(_disp, index=pivot.index, columns=pivot.columns)
        styled = display.style.apply(
            lambda _: pd.DataFrame(_css, index=pivot.index, columns=pivot.columns), axis=None)
    else:
        # 防御分支（当前所有调用方都是比率指标）：无统计行的简化渲染
        _pv = pivot.to_numpy(dtype=float)
        _rv = pivot_rsd.reindex(index=pivot.index, columns=pivot.columns).to_numpy(dtype=float)
        _disp = np.empty(_pv.shape, dtype=object)
        _css = np.empty(_pv.shape, dtype=object)
        for _i in range(_pv.shape[0]):
            for _j in range(_pv.shape[1]):
                _v = _pv[_i, _j]
                if np.isnan(_v):
                    _disp[_i, _j] = "-"
                    _css[_i, _j] = ""
                else:
                    _s = f"{_v:.2f}"
                    if show_rsd and not np.isnan(_rv[_i, _j]):
                        _s += f" (RSD:{_rv[_i, _j]:.1f}%)"
                    _disp[_i, _j] = _s
                    _css[_i, _j] = (f"background-color: {mpl.colors.to_hex(cmap_obj(norm(_v)))}"
                                    if color_on else "")
        display = pd.DataFrame(_disp, index=pivot.index, columns=pivot.columns)
        styled = display.style.apply(
            lambda _: pd.DataFrame(_css, index=pivot.index, columns=pivot.columns), axis=None)
    _df_kwargs = {}
    if full_height:
        # 高度 = 表头(双层列头按2行计) + 全部数据/统计行 → 无表内竖向滚动，页面滚动即看全部行
        _df_kwargs["height"] = 35 * (len(pivot) + 3) + 12
    try:
        st.dataframe(styled, width="stretch",
                     column_config={"物质": st.column_config.Column(width=170)},
                     **_df_kwargs)
    except Exception:
        # 兼容性兜底：Styler + column_config 组合异常时退回普通渲染
        st.dataframe(styled, width="stretch", **_df_kwargs)
    if metric_col in rate_cols:
        base_caption = f"两层列头：第一层=实验日期，第二层=组别名。{stat_desc}。"
    else:
        base_caption = "两层列头：第一层=实验日期，第二层=组别名。"
    if color_on and show_rsd:
        st.caption(base_caption + f"背景=数值颜色映射（{cmap_name}，色标 {vmin:.1f} ~ {vmax:.1f}），文本含 RSD。")
    elif color_on:
        st.caption(base_caption + f"背景=数值颜色映射（{cmap_name}，色标 {vmin:.1f} ~ {vmax:.1f}）。")
    elif show_rsd:
        st.caption(base_caption + "红色=RSD>15%。")
    else:
        st.caption(base_caption)


# ========================================================
# 6. 信息库
# ========================================================
def page_library():
    st.title("🗂️ 分析物 / SLM 信息库")
    if not REF_OK:
        st.warning("信息库模块未就绪。"); return

    # —— 信息库维护：上传替换 + 刷新 ——
    with st.expander("⬆️ 上传 / 刷新信息库（编辑药物、试剂信息用）"):
        st.caption("信息库文件位于 `data/analyte.xlsx`（分析物）与 `data/SLM.xlsx`（试剂）。"
                   "你可以在 Excel 中直接编辑它们（新增行/删减行/改性质），"
                   "或在这里上传替换文件；改动后点「🔄 刷新信息库」生效。")
        up1, up2 = st.columns(2)
        f_a = up1.file_uploader("上传分析物信息库（analyte.xlsx）", type=["xlsx"], key="lib_up_a")
        f_s = up2.file_uploader("上传 SLM 信息库（SLM.xlsx）", type=["xlsx"], key="lib_up_s")
        if f_a and st.button("💾 用上传文件替换 analyte.xlsx", key="lib_save_a"):
            try:
                pd.ExcelFile(f_a)  # 校验是有效 xlsx
                f_a.seek(0)
                config.ANALYTE_FILE.write_bytes(f_a.getvalue())
                st.success(f"已替换 `{config.ANALYTE_FILE.name}`，请点下方刷新生效。")
            except Exception as e:
                st.error(f"文件无效：{e}")
        if f_s and st.button("💾 用上传文件替换 SLM.xlsx", key="lib_save_s"):
            try:
                pd.ExcelFile(f_s)
                f_s.seek(0)
                config.SLM_FILE.write_bytes(f_s.getvalue())
                st.success(f"已替换 `{config.SLM_FILE.name}`，请点下方刷新生效。")
            except Exception as e:
                st.error(f"文件无效：{e}")
        if st.button("🔄 刷新信息库（重新读取 xlsx）", type="primary", key="lib_refresh"):
            try:
                R._load_analyte_chemical.cache_clear()
                R._analyte_basic_cached.cache_clear()
                R._analyte_info_cached.cache_clear()
                R.get_property_by_ph.cache_clear()
                R._name_forms_map.cache_clear()
                R._analyte_photo_cached.cache_clear()
                _analyte_numeric_cols_cached.cache_clear()
                try:
                    # 与试剂表重名的自定义标签自动转正（标识消失，信息按名称自然匹配）
                    _lib_now = R.get_slm_library()
                    S.purge_custom_reagents(set(_lib_now.get("main_names", []))
                                            | set(_lib_now.get("additive_names", [])))
                except Exception:
                    pass
                R._load_wide_table.cache_clear()
                R.list_ph_properties.cache_clear()
                for s in R.SLM_MAIN_SHEETS + R.SLM_ADDITIVE_SHEETS:
                    R._load_slm_sheet.cache_clear()
                _slm_cards_data.clear()
                st.success("✅ 信息库已刷新（新增/删减的物质与试剂即刻生效）。")
                st.rerun()
            except Exception as e:
                st.warning(f"刷新失败：{e}")
        st.caption("💡 提示：直接在 Excel 里改 `data/` 下的两个文件，改完回来点刷新即可——"
                   "新增列会自动识别（数值列可排序/散点，文本列可筛选）。")

        # —— 图片文件夹地址指定 ——
        st.markdown("**🖼️ 图片文件夹地址**（物质/试剂结构图所在文件夹，留空 = 用默认 `data/photo/` 下的目录）")
        pc1, pc2 = st.columns(2)
        cur_pa = S.get_photo_dir("analyte")
        cur_ps = S.get_photo_dir("slm")
        new_pa = pc1.text_input("分析物图片文件夹", value=str(cur_pa) if cur_pa else "", key="photo_dir_a",
                                placeholder=str(config.ANALYTE_PHOTO_DIR))
        new_ps = pc2.text_input("SLM 图片文件夹", value=str(cur_ps) if cur_ps else "", key="photo_dir_s",
                                placeholder=str(config.SLM_PHOTO_DIR))
        if st.button("💾 保存图片文件夹地址", key="photo_dir_save"):
            for kind, p in [("analyte", new_pa), ("slm", new_ps)]:
                p = p.strip()
                if p and not Path(p).is_dir():
                    st.warning(f"路径不存在（仍已保存，恢复存在后生效）：{p}")
                S.save_photo_dir(kind, p)
            st.success("✅ 已保存。物质/试剂详情将优先从自定义文件夹查找图片（按名称匹配 .png/.jpg）。")

    tab1, tab2 = st.tabs(["分析物", "SLM 试剂"])
    with tab1:
        df = R.get_analyte_basic()
        if df is not None and len(df):
            st.dataframe(df, width="stretch", hide_index=True)
        else:
            st.info("未加载到分析物信息")
    with tab2:
        _render_slm_library()


@st.cache_data(show_spinner=False)
def _slm_cards_data():
    if not REF_OK:
        return []
    out = []
    for sheet, label in [("LF-液膜类", "LF 主SLM"), ("DES-深共晶类", "DES 主SLM"),
                          ("IL-离子液体类", "IL 添加剂"), ("O-其他", "其他 添加剂")]:
        try:
            df = R._load_slm_sheet(sheet)
        except Exception:
            df = None
        if df is None:
            continue
        for name in df.index:
            out.append({"name": str(name), "category": label})
    return out


def _show_slm_card(name: str):
    info = R.get_slm_info(name)
    if info is None:
        st.warning(f"未找到 {name}"); return
    row = info.iloc[0]
    priority = S.get_display_columns("slm")
    have = [c for c in priority if c in info.columns and row[c] not in (None, "", "None")]
    cL, cR = st.columns([1, 2])
    with cL:
        img = R.get_slm_photo(name)
        if img:
            st.image(str(img), width="stretch")
        else:
            st.info("（无结构图）")
        st.markdown(f"**{name}**")
        for d in _slm_cards_data():
            if d["name"] == name:
                st.caption(d["category"]); break
    with cR:
        kv = {c: row[c] for c in have}
        if kv:
            st.markdown("**关键性质**")
            st.dataframe(pd.DataFrame(list(kv.items()), columns=["属性", "值"]),
                         width="stretch", hide_index=True)
    with st.expander("查看全部属性"):
        st.dataframe(info, width="stretch", hide_index=True)


def _render_slm_library():
    data = _slm_cards_data()
    if not data:
        st.info("未加载到 SLM 信息"); return
    cats = ["全部"] + sorted(set(d["category"] for d in data))
    sel_cat = st.selectbox("按分类筛选", cats)
    filtered = [d for d in data if sel_cat == "全部" or d["category"] == sel_cat]
    st.caption(f"共 {len(filtered)} 种试剂")
    kw = st.text_input("🔍 搜索试剂名", "")
    if kw:
        filtered = [d for d in filtered if kw.lower() in d["name"].lower()]
    selected = st.selectbox("选择试剂查看详情", ["（选择）"] + [d["name"] for d in filtered])
    if selected != "（选择）":
        _show_slm_card(selected)


# ========================================================
# 7. 设置页
# ========================================================
def build_voltage_df(current_sheets: list):
    """把选择的电流 sheet 记录（[{file, sheet}]）读取为电压表 DataFrame（多记录拼接，含来源列）"""
    if not current_sheets:
        return None
    frames = []
    for m in current_sheets:
        d = CUR.load_current_pair(m.get("file", ""), m.get("sheet", ""))
        if d:
            frames.append(pd.DataFrame({
                "时间(min)": d["time_min"],
                "电流值": d["current"],
                "来源sheet": [f"{m['file']}::{m['sheet']}"] * len(d["time_min"]),
            }))
    if not frames:
        return None
    return pd.concat(frames, ignore_index=True)


def page_calculator():
    st.title("🧪 实验前辅助计算器")
    st.caption("实验前的配制计算：稀释、标准需求量、流动相配比、摩尔浓度试剂配制。")

    tab1, tab2, tab3, tab4 = st.tabs(["💧 稀释计算", "🧴 标准需求量", "🧪 流动相配比", "⚖️ 摩尔浓度试剂配制"])

    # ---- Tab1: 稀释计算 ----
    with tab1:
        st.markdown("由储备液（初始浓度）配制目标浓度溶液：**C₁V₁ = C₂V₂×损失系数**")
        c1, c2, c3, c4 = st.columns(4)
        c1v = c1.number_input("初始浓度 C₁（储备液）", value=1000.0, format="%g", key="dil_c1")
        c2v = c2.number_input("需要的浓度 C₂", value=0.5, format="%g", key="dil_c2")
        v2 = c3.number_input("需要的体积 V₂", value=1.0, format="%g", key="dil_v2")
        k = c4.number_input("损失系数", value=1.0, min_value=1.0, step=0.05, key="dil_k",
                            help="最终体积 = 需要的体积 × 损失系数，避免实验损耗")
        v_total = v2 * k
        if c1v > 0 and v_total > 0:
            v1 = c2v * v_total / c1v
            v_dil = v_total - v1
            r1, r2, r3 = st.columns(3)
            r1.metric("加入储备液体积", f"{v1:.6g}")
            r2.metric("加入稀释液体积", f"{v_dil:.6g}")
            r3.metric("总体积", f"{v_total:.6g}")
            if v1 > v_total:
                st.warning("⚠️ 储备液体积大于总体积：C₂ 高于 C₁，无法稀释（应改为浓缩或降低 C₂）。")
        else:
            st.info("请输入初始浓度（>0）与需要的体积。")
        st.caption("单位自动保持一致（浓度同单位、体积同单位即可）。")

    # ---- Tab2: 标准需求量 ----
    with tab2:
        st.markdown("计算一次 EME 实验需要的标准溶液总量与储备液用量")
        a, b, c, d = st.columns(4)
        std_c = a.number_input("标准储备液浓度", value=10000.0, format="%g", key="std_c")
        std_c_unit = a.selectbox("储备液浓度单位", ["μg/L (≈ppb)", "mg/L"], key="std_c_unit",
                                 help="1 mg/L = 1000 μg/L；稀水溶液中 ppb ≈ μg/L")
        v_donor = b.number_input("供体相体积 (μL)", value=250.0, format="%g", key="std_vd")
        n_rep = c.number_input("每个SLM重复数", value=3, min_value=1, step=1, key="std_rep")
        c_donor = d.number_input("供体相浓度 (μg/L)", value=100.0, format="%g", key="std_cd")
        e, f_, g, h = st.columns(4)
        n_slm = e.number_input("实验测试的SLM个数", value=7, min_value=1, step=1, key="std_nslm")
        v_extra = f_.number_input("额外需要的标准体积 (μL)", value=250.0, format="%g", key="std_vx",
                                   help="如液质进样需要的标准溶液体积")
        v_acc = g.number_input("受体相体积 (μL)", value=50.0, format="%g", key="std_va")
        n_backup = h.number_input("备用的个数", value=3, min_value=0, step=1, key="std_nb")
        k2 = st.number_input("损失系数", value=1.0, min_value=1.0, step=0.05, key="std_k")
        import math as _math
        v_theory = n_rep * n_slm * v_donor
        v_rec = v_theory * k2 + n_backup * v_donor
        v_rec_up = _math.ceil(v_rec / 100.0) * 100  # 精确至百位并向上取整
        v_acc_min = n_rep * n_slm * v_acc
        if std_c > 0:
            c_std_ug = std_c * (1000.0 if std_c_unit == "mg/L" else 1.0)  # 统一换算为 μg/L
            v_stock = v_rec_up * c_donor / c_std_ug
            v_diluent = v_rec_up - v_stock
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("理论需要供体相体积 (μL)", f"{v_theory:.0f}")
            m2.metric("推荐供体相配制体积 (μL)", f"{v_rec_up:.0f}")
            m3.metric("受体相需要最小体积 (μL)", f"{v_acc_min:.0f}")
            m4.metric("受体相配制建议 (μL)", f"{_math.ceil((v_acc_min * k2 + n_backup * v_acc) / 100.0) * 100:.0f}")
            s1, s2 = st.columns(2)
            s1.metric("加入标准储备液 (μL)", f"{v_stock:.4g}")
            s2.metric("加入稀释液 (μL)", f"{v_diluent:.4g}")
            st.caption(f"计算：推荐体积 = 理论({v_theory:.0f})×损失系数({k2:g}) + 备用({n_backup})×{v_donor:g} = "
                       f"{v_rec:.1f} → 向上取整至百位 = {v_rec_up:.0f} μL；"
                       f"储备液用量 = {v_rec_up:.0f} × ({c_donor:g} μg/L ÷ {c_std_ug:.0f} μg/L) = {v_stock:.4g} μL。"
                       + (f" 额外标准体积 {v_extra:g} μL 已含在备用需求中，请按需并入。" if v_extra else ""))
        else:
            st.info("请输入标准储备液浓度（>0）。")

    # ---- Tab3: 流动相配比 ----
    with tab3:
        st.markdown("按比例计算流动相各组分加入量（甲酸先加、甲醇扣除其体积、水补足）")
        p1, p2, p3, p4 = st.columns(4)
        pct_water = p1.number_input("水比例 (%)", value=5.0, format="%g", key="mp_w")
        pct_meoh = p2.number_input("甲醇比例 (%)", value=95.0, format="%g", key="mp_m")
        pct_fa = p3.number_input("甲酸比例 (%)", value=0.1, format="%g", key="mp_f")
        vol_ml = p4.number_input("总体积 (mL)", value=500.0, format="%g", key="mp_v")
        total_pct = pct_water + pct_meoh + pct_fa
        if abs(total_pct - 100) > 0.01:
            st.warning(f"⚠️ 比例合计为 {total_pct:g}%（应为 100%）")
        v_fa = vol_ml * pct_fa / 100.0
        v_meoh = vol_ml * pct_meoh / 100.0 - v_fa
        v_water = vol_ml * pct_water / 100.0
        q1, q2, q3 = st.columns(3)
        q1.metric("甲酸加入量 (mL)", f"{v_fa:.4g}")
        q2.metric("甲醇加入量 (mL)", f"{v_meoh:.4g}")
        q3.metric("水加入量 (mL)", f"{v_water:.4g}")
        st.caption("说明：先量取甲酸，再量甲醇时扣除甲酸体积，最后加水至终体积。")

    # ---- Tab4: 摩尔浓度试剂配制 ----
    with tab4:
        st.markdown("按目标摩尔浓度配制试剂：由内置/自定义的**分子量+密度表**计算加入质量与体积")
        table = S.get_reagent_table()
        names = [r["名称"] for r in table]
        t1, t2, t3 = st.columns(3)
        sel_reagent = t3.selectbox("选择物质", names,
                                   index=names.index("甲酸") if "甲酸" in names else 0, key="mc_sel")
        c_mm = t1.number_input("需要的浓度 (mmol/L)", value=9.0, format="%g", key="mc_c")
        v_ml = t2.number_input("配制体积 (mL)", value=50.0, format="%g", key="mc_v")
        row = next((r for r in table if r["名称"] == sel_reagent), None)
        _mw_ok = bool(row) and row.get("分子量(g/mol)") is not None and row["分子量(g/mol)"] > 0
        # 密度可选：固体试剂无密度 → 只算质量，不算体积
        _dens_raw = row.get("密度(g/mL)") if row else None
        _dens_ok = _dens_raw is not None and _dens_raw > 0
        if _mw_ok:
            mw = row["分子量(g/mol)"]
            mass_raw = c_mm * v_ml * 0.001 * mw * 0.001  # mmol/L × mL → mol 换算
            mass = _math.ceil(mass_raw * 1000.0) / 1000.0  # 向上取整精确到千分位 (g)
            o1, o2, o3, o4 = st.columns(4)
            o1.metric("分子量 (g/mol)", f"{mw:g}")
            o2.metric("密度 (g/mL)", f"{_dens_raw:g}" if _dens_ok else "—")
            o3.metric("加入质量 (g)", f"{mass:.3f}")
            if _dens_ok:
                dens = _dens_raw
                vol_add = round(mass / dens, 3)
                o4.metric("加入体积 (mL)", f"{vol_add:.3f}")
                st.caption(f"计算：质量 = {c_mm:g} mmol/L × {v_ml:g} mL × 1e-6 L·mmol⁻¹·mL⁻¹ × {mw:g} g/mol "
                           f"= {mass_raw:.5f} g → 向上取整至千分位 = {mass:.3f} g；"
                           f"体积 = {mass:.3f} ÷ {dens:g} = {vol_add:.3f} mL（纯液体试剂用量筒/移液枪量取）。")
            else:
                o4.metric("加入体积 (mL)", "—")
                st.caption(f"计算：质量 = {c_mm:g} mmol/L × {v_ml:g} mL × 1e-6 L·mmol⁻¹·mL⁻¹ × {mw:g} g/mol "
                           f"= {mass_raw:.5f} g → 向上取整至千分位 = {mass:.3f} g。"
                           f"该物质未填密度（固体）：用天平称取质量即可，无需体积换算；"
                           f"如为液体可补填密度后自动算体积。")
        else:
            st.info("所选物质的分子量无效，请在下方试剂表中填写。")
        # —— 可编辑试剂表 ——
        st.markdown("---")
        st.markdown("**试剂表维护**（可直接编辑/新增行；保存后立即生效）")
        tbl = pd.DataFrame(table)
        edited_tbl = st.data_editor(tbl, num_rows="dynamic", key="reagent_tbl", width="stretch",
                                    hide_index=True,
                                    column_config={
                                        "分子量(g/mol)": st.column_config.NumberColumn(format="%.2f"),
                                        "密度(g/mL)": st.column_config.NumberColumn(format="%.3f"),
                                    })
        c_save1, c_save2 = st.columns(2)
        if c_save1.button("💾 保存试剂表", type="primary", key="reagent_save"):
            try:
                S.save_reagent_table(edited_tbl.to_dict("records"))
                st.success("✅ 试剂表已保存并生效。")
                st.rerun()
            except Exception as e:
                st.error(f"保存失败：{e}")
        if c_save2.button("↩️ 恢复内置默认表", key="reagent_reset"):
            S.save_reagent_table(S.DEFAULT_REAGENT_TABLE)
            st.session_state.pop("reagent_tbl", None)  # 丢弃编辑器旧控件状态，否则随后保存会把旧值写回
            st.success("已恢复内置默认。"); st.rerun()


def page_formulas():
    st.title("🧮 公式说明与自定义")
    st.caption("内置计算公式详解 + 自定义公式。自定义公式会在「实验详情 → 热图」中作为附加指标展示。")

    tab1, tab2, tab3 = st.tabs(["📐 内置公式说明", "✏️ 自定义公式", "🧪 公式测试器"])

    # ---- Tab1: 内置公式说明 ----
    with tab1:
        st.markdown("#### 1. 受体峰面积平均值 / RSD")
        st.markdown(
            "对每个实验组（如 1-A-1、1-A-2、1-A-3 三次重复）的**第1离子对峰面积**：\n\n"
            "- **平均值** = 三次重复峰面积的算术平均\n"
            "- **RSD%** = 标准差 ÷ 平均值 × 100（衡量重复性，**> 15% 标红**）\n\n"
            "数据来源：质谱导出的 ALLDATA.txt 中该组受体相（A）样本的峰面积列。")
        st.markdown("#### 2. 校正系数 corr")
        st.latex(r"corr = \frac{供体相标准品峰面积均值}{受体相标准品峰面积均值}")
        st.markdown(
            "**来源**：你在「新增实验」第2步输入的标准品名称，程序在数据中查找所有同名标准品进样取平均。\n\n"
            "**作用**：供体相与受体相溶液不同时排除基质效应；两相溶液相同时两者为同一样本，corr = 1。")
        st.markdown("#### 3. 回收率%")
        st.latex(r"回收率 = \frac{(受体峰面积均值 - 受体空白均值) \times Va}{(供体标准品均值 - 供体空白均值) \times Vd} \times corr \times 100\%")
        st.markdown(
            "| 变量 | 来源 |\n|---|---|\n"
            "| 受体峰面积均值 | 该组受体相（A）重复的峰面积平均（程序计算）|\n"
            "| 受体空白均值 | 你输入的受体空白名样本，取平均（程序查找）|\n"
            "| 供体标准品均值 | 你输入的供体标准品名样本，取平均（程序查找）|\n"
            "| 供体空白均值 | 你输入的供体空白名样本，取平均（程序查找）|\n"
            "| Va / Vd | 你在「新增实验」第3步输入的受体/供体相体积（μL，默认 50/250）|\n"
            "| corr | 上面的校正系数 |")
        st.markdown("#### 4. 供体相至SLM转移率%")
        st.latex(r"D \to M = 100\% - \frac{供体峰面积均值 - 供体空白均值}{供体标准品均值 - 供体空白均值} \times 100\%")
        st.markdown(
            "**含义**：萃取后供体相中剩余目标物的比例被扣除，得到从供体相进入膜的比例。\n\n"
            "**触发条件**：勾选「使用样品」且数据含供体相（S）样本时才计算。**超出 0~100 标红**。")
        st.markdown("#### 5. SLM至受体相转移率%")
        st.latex(r"M \to A = \frac{回收率}{D \to M} \times 100\%")
        st.markdown("**含义**：进入膜的目标物中最终到达受体相的比例。**超出 0~100 标红**。")
        st.markdown("#### 6. 标准品RSD%（质控用）")
        st.markdown(
            "同一次实验中，标准品多次进样（穿插在各组前后）峰面积的 RSD。用于质控看板："
            "**跨实验 ±15% 质控线** + 单次实验 RSD 柱状图（>15% 红）。")
        st.markdown("#### 7. 峰高 / 信噪比 / 半峰宽（辅助 QC 指标）")
        st.markdown(
            "| 指标 | 处理方式 | 展示位置 |\n|---|---|---|\n"
            "| **峰高（height）** | 与峰面积同类：每物质×组的受体相样品取**平均值 + RSD%**（>15% 标红）| 实验详情 → 热图区底部 |\n"
            "| **信噪比（SN）** | 与峰面积同类：从实验文件夹的 `*-SN.txt` 读取，受体相取**平均值 + RSD%** | 实验详情 → 热图区底部 |\n"
            "| **半峰宽（width50）** | 与保留时间同类：标准品取**平均值（不做RSD）** | 实验详情/汇总 + 质控看板半峰宽散点与跨实验 ±15% 质控图 |\n\n"
            "半峰宽均值写入数据库 `std_width_mean`（重新生成实验后有值）。")
        st.info("💡 所有公式均可在下方「自定义公式」中修改或另建，修改后只影响新展示的自定义指标，不覆盖内置结果。")

    # ---- Tab2: 自定义公式管理 ----
    with tab2:
        st.markdown("#### 可用变量")
        st.markdown(
            "| 变量名 | 含义 |\n|---|---|\n"
            "| `A_mean` | 该组受体相峰面积均值 |\n"
            "| `S_mean` | 该组供体相峰面积均值 |\n"
            "| `STD_acceptor_mean` | 受体相标准品峰面积均值 |\n"
            "| `STD_donor_mean` | 供体相标准品峰面积均值 |\n"
            "| `blk_acceptor_mean` | 受体相空白均值 |\n"
            "| `blk_donor_mean` | 供体相空白均值 |\n"
            "| `Va` / `Vd` | 受体/供体相体积（μL）|\n"
            "| `pH_donor` / `pH_acceptor` | 该组供体/受体相 pH |\n"
            "| `corr` | 校正系数 |")
        st.caption("支持 + - * / ( ) 和数字；示例：`(A_mean - blk_acceptor_mean) * Va / ((STD_donor_mean - blk_donor_mean) * Vd) * 100`")
        formulas = S.get_custom_formulas()
        new_list = []
        for i in range(S.MAX_CUSTOM_FORMULAS):
            cur = formulas[i] if i < len(formulas) else {"name": "", "expr": ""}
            with st.expander(f"自定义指标 {i+1}" + (f"：{cur['name']}" if cur["name"] else "（空）")):
                nm = st.text_input("指标名称（显示在热图标题）", value=cur.get("name", ""), key=f"cf_name_{i}")
                ex = st.text_area("表达式", value=cur.get("expr", ""), key=f"cf_expr_{i}", height=70)
                if nm and ex:
                    new_list.append({"name": nm, "expr": ex})
        c1, c2 = st.columns(2)
        if c1.button("💾 保存自定义公式", type="primary"):
            S.save_custom_formulas(new_list)
            st.success(f"已保存 {len(new_list)} 条自定义公式。到「实验详情 → 热图」查看自定义指标。")
        if c2.button("🗑 清空全部"):
            S.save_custom_formulas([])
            st.success("已清空。"); st.rerun()

    # ---- Tab3: 公式测试器 ----
    with tab3:
        st.markdown("输入表达式与一组测试值，实时查看计算结果。")
        texpr = st.text_input("表达式", value="(A_mean - blk_acceptor_mean) * Va / ((STD_donor_mean - blk_donor_mean) * Vd) * corr * 100", key="ft_expr")
        tc = st.columns(4)
        tvals = {
            "A_mean": tc[0].number_input("A_mean", value=1.0e6, format="%.3e", key="ft_A"),
            "S_mean": tc[1].number_input("S_mean", value=5.0e6, format="%.3e", key="ft_S"),
            "STD_donor_mean": tc[2].number_input("STD_donor_mean", value=2.0e7, format="%.3e", key="ft_sd"),
            "STD_acceptor_mean": tc[3].number_input("STD_acceptor_mean", value=2.0e7, format="%.3e", key="ft_sa"),
            "blk_acceptor_mean": 0.0, "blk_donor_mean": 0.0,
            "Va": tc[0].number_input("Va", value=50.0, key="ft_Va"),
            "Vd": tc[1].number_input("Vd", value=250.0, key="ft_Vd"),
            "pH_donor": tc[2].number_input("pH_donor", value=7.0, key="ft_phd"),
            "pH_acceptor": tc[3].number_input("pH_acceptor", value=3.0, key="ft_pha"),
            "corr": tc[0].number_input("corr", value=1.0, step=0.1, key="ft_corr"),
        }
        val, err = C.eval_custom_formula(texpr, tvals)
        if err:
            st.error(f"❌ {err}")
        else:
            st.success(f"✅ 计算结果 = **{val:.4f}**")


def page_settings():
    st.title("⚙️ 默认参数设置")
    st.caption("在此修改的默认值会保存到 `data/user_settings.json`，下次打开软件自动生效。")
    user_settings = S.load_settings()
    new_settings = {}
    for key, label, default, typ in S.SETTING_FIELDS:
        cur = user_settings.get(key, default)
        if typ == "number":
            try:
                cur = float(cur)
            except (ValueError, TypeError):
                cur = float(default)
            new_settings[key] = st.number_input(f"{label}", value=float(default) if cur is None else cur,
                                                key=f"set_{key}")
        else:
            new_settings[key] = st.text_input(f"{label}", value=str(default) if cur is None else str(cur),
                                              key=f"set_{key}")
    col1, col2 = st.columns(2)
    if col1.button("💾 保存设置", type="primary"):
        # 合并保存：只更新本页字段，不清掉配色偏好/参考线/自定义公式/试剂表等其他设置
        S.save_settings({**S.load_settings(), **new_settings})
        st.success("✅ 设置已保存，下次打开软件将使用这些默认值。")
    if col2.button("↩️ 恢复内置默认"):
        S.reset_settings()
        for _k in S.SETTING_FIELDS:  # 丢弃旧控件状态，否则随后点保存会把旧值写回
            st.session_state.pop(f"set_{_k[0]}", None)
        st.success("已恢复内置默认值。"); st.rerun()

    # ---- 自定义试剂标签管理 ----
    st.markdown("---")
    st.subheader("🧩 自定义试剂标签")
    st.caption("在新增实验的添加剂/试剂下拉框里**直接输入**的新名称会自动登记在这里，下拉框中以 🔧 标识；"
               "之后 SLM 试剂表若加入同名试剂，点「🔄 刷新信息库」后该标签自动转正（标识消失、信息按名称匹配）。")
    _lib_names_set = set()
    if REF_OK:
        try:
            _lib_m = R.get_slm_library()
            _lib_names_set = set(_lib_m.get("main_names", [])) | set(_lib_m.get("additive_names", []))
        except Exception:
            _lib_names_set = set()
    _cust_list = [n for n in S.get_custom_reagents() if n not in _lib_names_set]
    if not _cust_list:
        st.info("暂无自定义试剂标签。到「➕ 新增实验」的添加剂下拉框中直接输入新名称即可创建。")
    else:
        st.write(f"共 {len(_cust_list)} 个（仅存名称，不占用试剂表）：")
        for _ci in range(0, len(_cust_list), 4):
            _ccols = st.columns(4)
            for _cj, _n in enumerate(_cust_list[_ci:_ci + 4]):
                with _ccols[_cj]:
                    st.markdown(f"🔧 **{_n}**")
                    if st.button("🗑 删除", key=f"cust_del_{_n}"):
                        S.remove_custom_reagent(_n)
                        st.rerun()

    # ---- 样品名称格式（全局默认；单个实验可在新增实验页临时覆盖）----
    st.markdown("---")
    st.subheader("🔤 样品名称格式（全局默认）")
    st.caption("决定软件如何从文件里的样品名识别 组号 / 相（A=受体相，S=供体相）/ 重复号。"
               "分享给同门时各自在本机设置一次即可；单个实验也可在新增实验页临时覆盖（随实验保存）。")
    _preset_items = list(S.SAMPLE_FORMAT_PRESETS.items())
    _fmt_labels = [k for k, _ in _preset_items] + ["自定义正则"]
    _cur_fmt = S.get_sample_format()
    _cur_label = next((k for k, rx in _preset_items if rx == _cur_fmt), "自定义正则")
    def _sync_fmt_example_set():
        st.session_state["set_fmt_test"] = _sample_example_for(
            st.session_state.get("set_fmt_sel", ""))
    _fmt_label = st.radio("选择格式", _fmt_labels,
                          index=_fmt_labels.index(_cur_label), key="set_fmt_sel",
                          on_change=_sync_fmt_example_set)
    if _fmt_label == "自定义正则":
        _fmt_new = st.text_input("自定义正则（命名组：g=组号 ph=相位 r=重复号）",
                                 value=_cur_fmt, key="set_fmt_re").strip()
    else:
        _fmt_new = dict(_preset_items)[_fmt_label]
    if "set_fmt_test" not in st.session_state:
        st.session_state["set_fmt_test"] = _sample_example_for(_fmt_label)
    _test_name = st.text_input("样品名解析预览", key="set_fmt_test")
    try:
        re.compile(_fmt_new)
        _fmt_ok = {"g", "ph", "r"}.issubset(re.compile(_fmt_new).groupindex.keys())
        if not _fmt_ok:
            st.error("正则缺少命名组 g / ph / r，保存前请修正。")
        else:
            _pp_set = P.parse_sample_name(_test_name, pattern=_fmt_new)
            if _pp_set:
                st.caption(f"「{_test_name}」→ 组{_pp_set[0]} / {'受体相' if _pp_set[1] == 'A' else '供体相'} / 重复{_pp_set[2]}")
            else:
                st.caption(f"⚠️「{_test_name}」不符合该格式（将被当作质控/空白候选）——不影响保存，请确认格式本身正确。")
    except re.error as _fre:
        _fmt_ok = False
        st.error(f"正则无效：{_fre}")
    if st.button("💾 保存样品名称格式", key="set_fmt_save"):
        if not _fmt_ok:
            st.error("正则无效，未保存。")
        else:
            S.save_settings({**S.load_settings(), "SAMPLE_NAME_FORMAT": _fmt_new})
            st.success("✅ 已保存。之后的解析（含已有实验的重新生成，除非该实验保存过临时覆盖）都会按新格式识别样品名。")
            st.rerun()

    # ---- 统计参考线默认值（回收率 / 转移率分开）----
    st.markdown("---")
    st.subheader("📐 统计参考线默认值")
    st.caption("用于实验详情热图与实验汇总分表的按实验组统计。回收率与转移率各自独立设置；"
               "留空/清零表示该指标默认不启用参考线（页面内仍可临时启用并保存）。")
    rc_a, rc_b = st.columns(2)
    with rc_a:
        st.markdown("**回收率**")
        cur_r1a, cur_r2a = S.get_ref_lines("recovery")
        new_r1a = st.number_input("回收率 参考线1", value=float(cur_r1a if cur_r1a is not None else 20.0),
                                  key="set_ref_rec1")
        new_r2a = st.number_input("回收率 参考线2", value=float(cur_r2a if cur_r2a is not None else 40.0),
                                  key="set_ref_rec2")
        if st.button("💾 保存回收率参考线默认", key="set_ref_rec_save"):
            S.save_ref_lines(new_r1a if new_r1a and new_r1a > 0 else None,
                             new_r2a if new_r2a and new_r2a > 0 else None, kind="recovery")  # 清零=不启用（文档语义）
            st.success("已保存回收率参考线默认值。")
    with rc_b:
        st.markdown("**转移率（供体相至SLM / SLM至受体相 共用）**")
        cur_r1t, cur_r2t = S.get_ref_lines("transfer")
        new_r1t = st.number_input("转移率 参考线1（0=不启用）",
                                  value=float(cur_r1t if cur_r1t is not None else 0.0),
                                  key="set_ref_tr1")
        new_r2t = st.number_input("转移率 参考线2（0=不启用）",
                                  value=float(cur_r2t if cur_r2t is not None else 0.0),
                                  key="set_ref_tr2")
        if st.button("💾 保存转移率参考线默认", key="set_ref_tr_save"):
            S.save_ref_lines(new_r1t if new_r1t > 0 else None,
                             new_r2t if new_r2t > 0 else None,
                             kind="transfer")
            st.success("已保存转移率参考线默认值（0=默认不启用）。")

    # ---- 详情展示列自定义 ----
    st.markdown("---")
    st.subheader("🎛️ 物质 / SLM 详情展示列自定义")
    st.caption("勾选要在「物质详情」和「SLM 详情」中展示的分子描述符列。")
    if REF_OK:
        tab_a, tab_s = st.tabs(["分析物展示列", "SLM 展示列"])
        with tab_a:
            ab = R.get_analyte_basic()
            if ab is not None and len(ab):
                all_cols = [c for c in ab.columns if c not in ("name", "Englishname")]
                cur_a = S.get_display_columns("analyte")
                sel_a = st.multiselect("分析物展示列", all_cols, default=[c for c in cur_a if c in all_cols],
                                       key="disp_a")
                if st.button("保存分析物展示列"):
                    S.save_display_columns("analyte", sel_a)
                    st.success("已保存分析物展示列。")
            else:
                st.info("未加载到分析物信息。")
        with tab_s:
            slm_all = R.get_slm_all()
            if slm_all is not None and len(slm_all):
                all_cols_s = [c for c in slm_all.columns if c not in ("试剂名称", "_sheet")]
                cur_s = S.get_display_columns("slm")
                sel_s = st.multiselect("SLM 展示列", all_cols_s, default=[c for c in cur_s if c in all_cols_s],
                                       key="disp_s")
                if st.button("保存SLM展示列"):
                    S.save_display_columns("slm", sel_s)
                    st.success("已保存 SLM 展示列。")
            else:
                st.info("未加载到 SLM 信息。")
    else:
        st.info("信息库模块未就绪，无法列出可选列。")


# ---------- 数据观察记录 ----------
def render_global_observation_sidebar():
    """主导航区侧栏的全局观察记录入口（固定在功能导航下方）。
    若在「📋 实验详情」选中了实验，自动关联该实验；可勾选改为整体记录。
    """
    # 清空输入必须在 text_area 实例化【之前】写其 key（实例化后再写会 StreamlitAPIException）
    if st.session_state.pop("_obs_clear_txt", False):
        st.session_state["obs_txt_global"] = ""
    with st.sidebar:
        st.markdown("---")
        with st.expander("📝 数据观察记录"):
            # 实验详情页下拉的 key=el_sel：widget 最新值在本轮脚本运行前就已写入
            # session_state（_cur_detail_folder 由页面函数设置，晚于侧栏渲染，会滞后一轮）
            folder = st.session_state.get("el_sel") or st.session_state.get("_cur_detail_folder")
            attach_folder = None
            _folder_exists = False
            if folder:
                try:
                    _folder_exists = get_db().execute(
                        "SELECT 1 FROM experiments WHERE folder=?", (folder,)).fetchone() is not None
                except Exception:
                    _folder_exists = False
            if folder and _folder_exists:
                st.caption(f"将关联当前选中实验：{folder}")
                only_global = st.checkbox("改为整体记录（不关联实验）", key="obs_global_only")
                attach_folder = None if only_global else folder
            else:
                st.caption("未选中实验 → 保存为整体记录（在「📋 实验详情」选中实验后自动关联）。")
            txt = st.text_area("记录内容", value="", key="obs_txt_global",
                               placeholder="实验现象、异常、备忘…")
            obs_img = None
            if attach_folder:
                obs_img = st.file_uploader("📷 附实验图片（可选，自动压缩）",
                                           type=["png", "jpg", "jpeg", "bmp", "webp"],
                                           key="obs_img_upload")
            if st.button("💾 保存记录", key="obs_save_global"):
                if txt.strip():
                    saved = False
                    if attach_folder:
                        conn = get_db()
                        row = conn.execute("SELECT id, exp_date FROM experiments WHERE folder=?",
                                           (attach_folder,)).fetchone()
                        if row:
                            _img_rel = ""
                            if obs_img is not None:
                                try:
                                    _img_rel = M.save_observation_image(attach_folder,
                                                                        row["exp_date"] or "", obs_img)
                                except Exception as _ie:
                                    st.warning(f"图片保存失败（记录仍会保存）：{_ie}")
                            M.add_observation("experiment", txt.strip(),
                                              experiment_id=int(row["id"]),
                                              exp_date=row["exp_date"], folder=attach_folder,
                                              image=_img_rel)
                            saved = True
                    if not saved:
                        M.add_observation("global", txt.strip())
                    st.session_state["_obs_clear_txt"] = True  # 下一轮渲染前清空输入框
                    st.success("✅ 已保存（可在「📝 观察记录」页查看/导出全部记录）。")
                else:
                    st.warning("请先输入内容。")


def page_observations():
    st.title("📝 数据观察记录")
    st.caption("整体观察与单次实验观察的汇总页，可导出 CSV 方便整理。")
    conn = get_db()
    df = M.list_observations()
    if len(df) == 0:
        st.info("暂无观察记录。在各功能页侧栏的「📝 数据观察记录」中添加。")
        return
    show = df.copy()
    show["scope"] = show["scope"].map({"global": "🌐 整体", "experiment": "🔬 单次实验"}).fillna(show["scope"])
    show = show.rename(columns={"scope": "类型", "exp_date": "实验日期", "folder": "实验文件夹",
                                "text": "记录内容", "created_at": "记录时间"})
    st.dataframe(show[["类型", "实验日期", "实验文件夹", "记录内容", "记录时间"]],
                 width="stretch", hide_index=True)
    # 图片展示（按记录逐条；图片存在实验文件夹里，自动命名）
    _with_img = df[df["image"].notna() & (df["image"].astype(str).str.len() > 0)] if "image" in df.columns else df.iloc[0:0]
    if len(_with_img):
        st.markdown("**📷 附图记录**")
        for _, _r in _with_img.iterrows():
            with st.expander(f"{_r['created_at']} | {_r['folder'] or '整体'} | {_r['text'][:30]}"):
                _p = M.get_observation_image_path(str(_r["image"]))
                if _p:
                    st.image(str(_p), width="stretch")
                    st.caption(f"文件：{_r['image']}")
                else:
                    st.caption(f"（图片文件缺失：{_r['image']}）")
    st.download_button("⬇️ 导出 CSV", data=show.to_csv(index=False).encode("utf-8-sig"),
                       file_name="观察记录.csv", mime="text/csv")
    # 逐条修改（保存后实验详情页的观察记录同步更新——同一数据库）
    st.markdown("---")
    st.markdown("**✏️ 修改记录**（修改后「📋 实验详情」中该实验的观察记录同步更新）")
    with st.expander("展开修改某条记录"):
        _ed_idx = st.selectbox("选择要修改的记录", show.index.tolist(),
                               format_func=lambda i: f"{show.loc[i, '记录时间']} | {show.loc[i, '类型']} | {str(show.loc[i, '记录内容'])[:40]}")
        _ed_text = st.text_area("记录内容（修改）", value=str(show.loc[_ed_idx, "记录内容"]),
                                key=f"obs_edit_{int(show.loc[_ed_idx, 'id'])}")
        if st.button("💾 保存修改", key="obs_edit_save"):
            if _ed_text.strip():
                M.update_observation(int(show.loc[_ed_idx, "id"]), _ed_text.strip())
                st.session_state["_flash_msg"] = "已修改该条观察记录（详情页同步更新）。"
                st.rerun()
            else:
                st.warning("内容不能为空。")
    st.markdown("---")
    with st.expander("➕ 新增整体观察记录"):
        txt = st.text_area("记录内容", value="", key="obs_page_txt")
        if st.button("💾 保存", key="obs_page_save"):
            if txt.strip():
                M.add_observation("global", txt.strip())
                st.success("已保存。"); st.rerun()
    with st.expander("🗑 删除记录"):
        _del_idx = st.selectbox("选择要删除的记录", show.index.tolist(),
                                format_func=lambda i: f"{show.loc[i, '记录时间']} | {str(show.loc[i, '记录内容'])[:40]}（#{int(show.loc[i, 'id'])}）")
        if st.button("确认删除", type="secondary"):
            M.delete_observation(int(show.loc[_del_idx, "id"]))
            st.session_state["_flash_msg"] = "已删除该条观察记录。"
            st.rerun()


# ---------- 主导航 ----------
st.sidebar.title("🧪 EME 实验数据管理")
menu = st.sidebar.radio(
    "功能导航",
    ["🧪 计算器", "➕ 新增实验", "📋 实验详情", "📈 实验汇总", "📊 质控看板", "🔧 自定义看板",
     "📝 观察记录", "🧮 公式说明", "🗂️ 信息库", "⚙️ 设置"],
    index=0)
render_global_observation_sidebar()
if st.session_state.get("_flash_msg"):
    st.success(st.session_state.pop("_flash_msg"))

if menu == "➕ 新增实验":
    page_new_experiment()
elif menu == "📋 实验详情":
    page_experiment_list()
elif menu == "📈 实验汇总":
    page_summary()
elif menu == "📊 质控看板":
    page_qc_dashboard()
elif menu == "🔧 自定义看板":
    page_custom_dashboard()
elif menu == "🧪 计算器":
    page_calculator()
elif menu == "📝 观察记录":
    page_observations()
elif menu == "🧮 公式说明":
    page_formulas()
elif menu == "🗂️ 信息库":
    page_library()
elif menu == "⚙️ 设置":
    page_settings()

# 侧栏末尾：数据目录信息
st.sidebar.caption(f"数据目录：`{config.DATA_DIR}`")
st.sidebar.caption("💡 可在新增实验页直接拖入 txt 数据；公式自定义见「🧮 公式说明」")
