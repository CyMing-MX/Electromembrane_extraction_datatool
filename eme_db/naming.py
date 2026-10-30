# -*- coding: utf-8 -*-
"""
EME 实验数据管理系统 - 组别自动命名（v2）

按实验目的生成组别名（前缀-段1-段2-段3[-电压-时间]）：
  前缀：测试SLM→SLM、测试供体相→D、测试受体相→A、
        测试电压→V、测试时间→T、综合优化（多因素）→Mul
  段1 = 主SLM（混合膜用 + 连接）+（SLM添加剂，如有；多个用 + 连接）
  段2 = 供体相成分及pH（如 AceA2.4）+（供体相添加剂，如有）
  段3 = 受体相成分及pH（如 FA3.0）+（受体相添加剂，如有）
  电压/时间/综合优化在末尾追加 -50V-30min（缺失的部分省略）

示例：NPPE+DEHP 膜 / AceA pH2.4+DEHP 供体 / FA pH3.0 受体 / 50V 30min
  测试SLM：SLM-NPPE+DEHP-AceA2.4+DEHP-FA3.0
  测试供体相：D-NPPE+DEHP-AceA2.4+DEHP-FA3.0
  测试受体相：A-NPPE+DEHP-AceA2.4+DEHP-FA3.0
  测试电压：V-NPPE+DEHP-AceA2.4+DEHP-FA3.0-50V-30min
  测试时间：T-NPPE+DEHP-AceA2.4+DEHP-FA3.0-50V-30min
  综合优化：Mul-NPPE+DEHP-AceA2.4+DEHP-FA3.0-50V-30min

名称里的比例后缀（"DEHP (5%)"）在命名中剥离，只保留物质名。
"""
from __future__ import annotations
import json
import re
from typing import Optional

# 实验目的 key（与 app.py 的 EXP_TYPES 一致）
EXP_TYPE_KEYS = {"slm", "donor", "acceptor", "time", "voltage", "all", "custom"}

# 实验目的 → 前缀
EXP_TYPE_PREFIX = {
    "slm": "SLM",
    "donor": "D",
    "acceptor": "A",
    "voltage": "V",
    "time": "T",
    "all": "Mul",
}

_RATIO_RE_CACHE = {}


def _ratio_re():
    import re as _re
    if "rx" not in _RATIO_RE_CACHE:
        _RATIO_RE_CACHE["rx"] = _re.compile(
            r"\s*[(（]\d+(?:\.\d+)?\s*(?:%|mM|μM|uM|µM|M|mg/L|g/L|μg/L|ug/L)?\s*[)）]\s*$",
            _re.IGNORECASE)
    return _RATIO_RE_CACHE["rx"]


def _clean(name) -> str:
    """剥比例后缀；混合主SLM "NPPE (80%) + TBP (20%)" → "NPPE+TBP" """
    if name is None:
        return ""
    s = str(name)
    if "+" in s:  # 混合物：逐段剥比例后用 + 连接
        parts = [_ratio_re().sub("", p).strip() for p in s.split("+")]
        return "+".join([p for p in parts if p])
    return _ratio_re().sub("", s).strip()


def _fmt_additives(additives_raw) -> list:
    """添加剂列表 → 剥比例后的物质名列表。支持 JSON 字符串或 list"""
    if additives_raw is None:
        return []
    if isinstance(additives_raw, str):
        try:
            lst = json.loads(additives_raw)
        except Exception:
            lst = [additives_raw] if additives_raw else []
    else:
        lst = list(additives_raw)
    out = []
    for x in lst:
        c = _clean(x)
        if c:
            out.append(c)
    return out


def _fmt_ph(ph) -> str:
    """pH 数值 → 紧凑文本：2.4→"2.4"，3.0→"3.0"（保留 .0 与示例一致）"""
    if ph is None:
        return ""
    try:
        s = f"{float(ph):g}"
    except (TypeError, ValueError):
        return ""
    return s if ("." in s or "e" in s or "E" in s) else s + ".0"


def _phase_seg(composition, ph, additives) -> str:
    """相段：成分名+pH（如 AceA2.4）+（添加剂，如有，+ 连接）"""
    seg = f"{str(composition or '').strip()}{_fmt_ph(ph)}".strip()
    adds = _fmt_additives(additives)
    if adds:
        seg = f"{seg}+{'+'.join(adds)}" if seg else "+".join(adds)
    return seg


def generate_group_label(
    exp_type: str,
    slm_main: str = "",
    slm_additives=None,
    pH_donor: Optional[float] = None,
    donor_additives=None,
    pH_acceptor: Optional[float] = None,
    acceptor_additives=None,
    extract_time: Optional[float] = None,
    extract_voltage: Optional[float] = None,
    donor_composition: str = "",
    acceptor_composition: str = "",
) -> str:
    """
    按实验目的规则生成组别名（v2 格式，见模块 docstring）。
    exp_type: slm/donor/acceptor/time/voltage/all/custom（custom/未知 → 返回空，手动命名）
    """
    prefix = EXP_TYPE_PREFIX.get(exp_type)
    if prefix is None:
        return ""

    slm_seg = _clean(slm_main)
    if not slm_seg:
        # 没有 SLM 信息，无法生成有意义的别名（调用方用 M组号 兜底）
        return ""
    slm_adds = _fmt_additives(slm_additives)
    if slm_adds:
        slm_seg = f"{slm_seg}+{'+'.join(slm_adds)}"

    parts = [prefix, slm_seg,
             _phase_seg(donor_composition, pH_donor, donor_additives),
             _phase_seg(acceptor_composition, pH_acceptor, acceptor_additives)]
    label = "-".join([p for p in parts if p])

    # 电压 / 时间 / 综合优化：末尾追加电压和时间（缺失的省略）
    if exp_type in ("voltage", "time", "all"):
        tail = []
        if extract_voltage is not None:
            try:
                tail.append(f"{float(extract_voltage):g}V")
            except (TypeError, ValueError):
                pass
        if extract_time is not None:
            try:
                tail.append(f"{float(extract_time):g}min")
            except (TypeError, ValueError):
                pass
        if tail:
            label = f"{label}-{'-'.join(tail)}" if label else "-".join(tail)
    return label


def build_group_labels(exp_type: str, base_cond: dict, group_conds: dict) -> dict:
    """
    批量生成组别名。
    返回 {组号: 组别名}
    base_cond, group_conds: 见 app.py 录入（每组的条件已合并基准）
    """
    labels = {}
    for grp, cond in group_conds.items():
        # cond 应已合并基准，可直接用
        lbl = generate_group_label(
            exp_type,
            slm_main=cond.get("slm_main") or base_cond.get("slm_main", ""),
            slm_additives=cond.get("slm_additives", base_cond.get("slm_additives")),
            pH_donor=cond.get("pH_donor", base_cond.get("pH_donor")),
            donor_additives=cond.get("donor_additives", base_cond.get("donor_additives")),
            pH_acceptor=cond.get("pH_acceptor", base_cond.get("pH_acceptor")),
            acceptor_additives=cond.get("acceptor_additives", base_cond.get("acceptor_additives")),
            extract_time=cond.get("extract_time", base_cond.get("extract_time")),
            extract_voltage=cond.get("extract_voltage", base_cond.get("extract_voltage")),
            donor_composition=cond.get("donor_composition", base_cond.get("donor_composition", "")),
            acceptor_composition=cond.get("acceptor_composition", base_cond.get("acceptor_composition", "")),
        )
        labels[grp] = lbl or f"M{grp}"
    return labels


if __name__ == "__main__":
    # 自检：用户给定示例
    kw = dict(slm_main="NPPE", slm_additives=["DEHP (5%)"],
              pH_donor=2.4, donor_composition="AceA", donor_additives=["DEHP (3%)"],
              pH_acceptor=3.0, acceptor_composition="FA", acceptor_additives=[],
              extract_voltage=50, extract_time=30)
    assert generate_group_label("slm", **kw) == "SLM-NPPE+DEHP-AceA2.4+DEHP-FA3.0", generate_group_label("slm", **kw)
    assert generate_group_label("donor", **kw) == "D-NPPE+DEHP-AceA2.4+DEHP-FA3.0"
    assert generate_group_label("acceptor", **kw) == "A-NPPE+DEHP-AceA2.4+DEHP-FA3.0"
    assert generate_group_label("voltage", **kw) == "V-NPPE+DEHP-AceA2.4+DEHP-FA3.0-50V-30min"
    assert generate_group_label("time", **kw) == "T-NPPE+DEHP-AceA2.4+DEHP-FA3.0-50V-30min"
    assert generate_group_label("all", **kw) == "Mul-NPPE+DEHP-AceA2.4+DEHP-FA3.0-50V-30min"
    # 无添加剂 / 混合膜 / 缺失条件
    assert generate_group_label("slm", slm_main="NPPE", pH_acceptor=3.0,
                                acceptor_composition="FA") == "SLM-NPPE-FA3.0"
    assert generate_group_label("slm", slm_main="NPPE (80%) + TBP (20%)",
                                pH_acceptor=7, acceptor_composition="FA") == "SLM-NPPE+TBP-FA7.0"
    assert generate_group_label("custom", slm_main="NPPE") == ""
    assert generate_group_label("slm", slm_main="") == ""
    print("命名规则自检全部通过")
    print(generate_group_label("slm", **kw))
    print(generate_group_label("all", **kw))
