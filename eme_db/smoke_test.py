# -*- coding: utf-8 -*-
"""全页面冒烟测试 + 本次修复点的针对性验证。运行: python smoke_test.py"""
import sys, os, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import models as M
import settings as S
import calculator as C
import parser as P
import excel_builder as EB
import pandas as pd

from streamlit.testing.v1 import AppTest

OK, FAIL = [], []


def check(name, fn):
    try:
        fn()
        OK.append(name)
        print(f"  ✅ {name}")
    except Exception as e:
        FAIL.append((name, repr(e)))
        print(f"  ❌ {name}: {e!r}")


def main():
    print("== 1. 单元级验证 ==")

    # 1a. excluded_samples dict 格式兼容
    def t_excl():
        conn = M.get_conn()
        # 找一个实验来演练；用内存级插入避免破坏：直接用最大 id+1 不行，找真实实验
        row = conn.execute("SELECT id FROM experiments LIMIT 1").fetchone()
        if row is None:
            print("    （无实验记录，跳过）"); return
        eid = int(row["id"])
        old_raw = M._load_excluded_raw(conn, eid)  # 测试前的完整 excluded_samples（含 std 键）
        try:
            M.set_excluded_samples(conn, eid, {"1-A-1"})
            assert M.get_excluded_samples(conn, eid) == {"1-A-1"}
            M.set_excluded_std(conn, eid, std_acceptor={("STD-X", 2)}, std_donor={("STD-Y", 1)})
            assert M.get_excluded_samples(conn, eid) == {"1-A-1"}, "样本剔除被标准品写入冲掉"
            st = M.get_excluded_std(conn, eid)
            assert st["std_acceptor"] == {("STD-X", 2)} and st["std_donor"] == {("STD-Y", 1)}
            M.set_excluded_std(conn, eid, std_acceptor={("STD-X", 3)})
            assert M.get_excluded_std(conn, eid)["std_acceptor"] == {("STD-X", 3)}
            assert M.get_excluded_std(conn, eid)["std_donor"] == {("STD-Y", 1)}
            M.set_excluded_samples(conn, eid, set())
            assert M.get_excluded_std(conn, eid)["std_donor"] == {("STD-Y", 1)}, "标准品剔除被样本写入冲掉"
        finally:
            # 原样还原测试前的完整 excluded_samples（samples 与 std 键都不丢不添）
            conn.execute("UPDATE experiments SET excluded_samples=? WHERE id=?",
                         (json.dumps(old_raw, ensure_ascii=False) if old_raw else None, eid))
            conn.commit()
    check("excluded_samples dict 兼容 + 互不冲掉", t_excl)

    # 1b. 配色偏好
    def t_cmap():
        S.save_cmap_pref("Blues")
        assert S.get_cmap_pref() == "Blues"
        S.save_cmap_pref("不存在的")
        assert S.get_cmap_pref() == "Blues", "非法值应回退"
        S.save_cmap_pref("RdYlGn")
    check("配色偏好保存/读取/非法回退", t_cmap)

    # 1c. Excel 生成含 SN/峰高/半峰宽表
    def t_excel():
        import config
        exps_dir = config.EXPERIMENTS_DIR
        folders = sorted([f for f in exps_dir.iterdir() if f.is_dir()]) if exps_dir.exists() else []
        if not folders:
            print("    （无实验文件夹，跳过）"); return
        res = P.parse_experiment(folders[-1])
        out = os.path.join(os.environ.get("TEMP", "."), "_smoke_DATA.xlsx")
        EB.build_excel(out, res, {"exp_date": res.file_date or "test"}, 50.0, 250.0, True,
                       gcond=None)
        from openpyxl import load_workbook
        wb = load_workbook(out)
        names = wb.sheetnames
        for need in ["总表", "数据汇总", "RT表", "电压表", "M1"]:
            assert need in names, f"缺 {need}"
        print(f"    sheets: {names}")
        os.remove(out)
    check("Excel 生成（含新 SN/峰高/半峰宽表按需）", t_excel)

    print("== 2. 页面冒烟（AppTest）==")
    PAGES = ["🧪 计算器", "➕ 新增实验", "📋 实验详情", "📈 实验汇总", "📊 质控看板",
             "🔧 自定义看板", "📝 观察记录", "🧮 公式说明", "🗂️ 信息库", "⚙️ 设置"]


    def run_page(menu):
        at = AppTest.from_file("app.py", default_timeout=120)
        at.run()
        assert not at.exception, f"默认页异常: {at.exception}"
        # 切换到目标页
        at.sidebar.radio[0].set_value(menu).run()
        assert not at.exception, f"{menu} 页异常"
        return at


    for pg in PAGES:
        check(f"页面 {pg}", lambda p=pg: run_page(p))

    # 2b. 质控看板分区切换（fragment 分支全覆盖）
    def t_qc_tabs():
        at = AppTest.from_file("app.py", default_timeout=180)
        at.run()
        at.sidebar.radio[0].set_value("📊 质控看板").run()
        assert not at.exception
        # 找到分区 radio（qc_tab_sel）
        radios = [r for r in at.radio if r.key == "qc_tab_sel"]
        assert radios, "找不到分区 radio"
        for opt in radios[0].options:
            try:
                radios[0].set_value(opt).run()
            except KeyError as e:
                # AppTest 框架局限：fragment 分支切换后收集过期控件状态时报 KeyError，
                # 真实浏览器不经过这条代码路径（浏览器只提交真实存在的控件状态）
                if "st.session_state has no key" in str(e):
                    print(f"    （AppTest 已知局限，忽略：{opt}）")
                    at = AppTest.from_file("app.py", default_timeout=180)
                    at.run()
                    at.sidebar.radio[0].set_value("📊 质控看板").run()
                    radios = [r for r in at.radio if r.key == "qc_tab_sel"]
                    radios[0].set_value(opt).run()
                else:
                    raise
            assert not at.exception, f"分区 {opt} 异常"
    check("质控看板分区逐一渲染", t_qc_tabs)

    # 2c. 实验汇总：指标切换 + 展开全部药物行开关
    def t_sum_fs():
        at = AppTest.from_file("app.py", default_timeout=180)
        at.run()
        at.sidebar.radio[0].set_value("📈 实验汇总").run()
        assert not at.exception
        mradio = [r for r in at.radio if r.key == "sum_metric_sel"]
        assert mradio, "找不到指标 radio"
        for opt in mradio[0].options:
            try:
                mradio[0].set_value(opt).run()
            except KeyError as e:
                if "st.session_state has no key" not in str(e):
                    raise
                # AppTest fragment 局限，重建会话后重试
                at = AppTest.from_file("app.py", default_timeout=180)
                at.run()
                at.sidebar.radio[0].set_value("📈 实验汇总").run()
                mradio = [r for r in at.radio if r.key == "sum_metric_sel"]
                mradio[0].set_value(opt).run()
            assert not at.exception, f"指标 {opt} 异常"
        # 展开全部药物行开关（切换两种状态）
        cb = [c for c in at.checkbox if c.key == "sum_expand_all"]
        assert cb, "找不到 sum_expand_all 开关"
        cb[0].uncheck().run()
        assert not at.exception
        cb[0].check().run()
        assert not at.exception
    check("实验汇总 指标切换+展开开关", t_sum_fs)

    # 2d. 配色偏好经 UI 保存
    def t_cmap_ui():
        at = AppTest.from_file("app.py", default_timeout=120)
        at.run()
        at.sidebar.radio[0].set_value("📈 实验汇总").run()
        sels = [s for s in at.selectbox if s.key == "sum_cmap"]
        if sels:
            sels[0].set_value("Blues").run()
            assert not at.exception
            assert S.get_cmap_pref() == "Blues", "选择配色未持久化"
            S.save_cmap_pref("RdYlGn")  # 还原
    check("汇总页配色选择自动持久化", t_cmap_ui)

    print()
    print(f"通过 {len(OK)} / {len(OK)+len(FAIL)}")
    if FAIL:
        print("失败项：")
        for n, e in FAIL:
            print(f"  - {n}: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
