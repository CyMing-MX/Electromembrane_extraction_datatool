# -*- coding: utf-8 -*-
"""
EME 实验数据管理系统 - exe 启动器（PyInstaller 入口）

行为：
1. 定位打包在 _internal/eme_db/app.py 的应用
2. 找一个空闲端口（8501 起）
3. 后台线程探活端口就绪后自动打开浏览器
4. 前台运行 Streamlit（关闭窗口/Ctrl+C 即停止服务）
"""
import os
import socket
import sys
import threading
import time
import webbrowser


def base_dir() -> str:
    """exe 所在目录（用户数据 data/ 放这里）"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def resource_dir() -> str:
    """打包资源目录（_internal，含 eme_db/app.py）"""
    if getattr(sys, "frozen", False):
        return sys._MEIPASS  # onedir 模式指向 exe旁的 _internal
    return os.path.dirname(os.path.abspath(__file__))


def find_free_port(start: int = 8501, tries: int = 20) -> int:
    for p in range(start, start + tries):
        s = socket.socket()
        try:
            s.bind(("127.0.0.1", p))
            s.close()
            return p
        except OSError:
            s.close()
    return start


def open_browser_when_ready(port: int, delay: float = 0.0):
    if delay:
        time.sleep(delay)
    for _ in range(120):  # 最多等 120 秒（打包版首次启动较慢）
        s = socket.socket()
        s.settimeout(1)
        rc = s.connect_ex(("127.0.0.1", port))
        s.close()
        if rc == 0:
            webbrowser.open(f"http://localhost:{port}")
            return
        time.sleep(1)


def selftest() -> int:
    """打包完整性自检：EME_Data.exe selftest"""
    ok, fail = [], []
    mods = ["streamlit", "pandas", "pyarrow", "openpyxl", "plotly", "matplotlib",
            "simpleeval", "altair"]
    for m in mods:
        try:
            __import__(m)
            ok.append(m)
        except Exception as e:
            fail.append(f"{m}: {type(e).__name__}: {e}")
    try:
        import pandas as pd
        import pyarrow as pa
        from pyarrow.vendored.version import Version  # 常见收集遗漏点
        tbl = pa.Table.from_pandas(pd.DataFrame({"a": [1]}))
        ok.append(f"pa.Table.from_pandas(rows={tbl.num_rows})")
    except Exception as e:
        fail.append(f"pyarrow.Table: {type(e).__name__}: {e}")
    try:
        import sys as _s
        base = os.path.dirname(os.path.abspath(__file__))
        if getattr(_s, "frozen", False):
            base = os.path.join(_s._MEIPASS, "eme_db")
        assert os.path.exists(os.path.join(base, "app.py")), "app.py 不在包内"
        ok.append("eme_db/app.py")
    except Exception as e:
        fail.append(f"资源: {e}")
    print("SELFTEST OK: " + ", ".join(ok) if not fail else "SELFTEST FAIL:\n  " + "\n  ".join(fail))
    return 0 if not fail else 1


def main():
    if len(sys.argv) > 1 and sys.argv[1] == "selftest":
        return selftest()
    port = find_free_port(8501)
    app_path = os.path.join(resource_dir(), "eme_db", "app.py")
    if not os.path.exists(app_path):
        # 兜底：exe 旁直接放 eme_db 的场景
        alt = os.path.join(base_dir(), "eme_db", "app.py")
        if os.path.exists(alt):
            app_path = alt
        else:
            print(f"[ERROR] 找不到应用文件：{app_path}")
            input("按回车退出...")
            return 1

    # 用户数据目录（config.py 会自动发现 exe 旁的 data/）
    os.environ.setdefault("EME_PORT", str(port))
    # frozen 环境下 Streamlit 会自动进入 developmentMode（禁止自定义端口）——强制关闭
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"
    # matplotlib 用纯软件渲染（Agg），避免打包进几十 MB 的 Qt 后端
    os.environ.setdefault("MPLBACKEND", "Agg")

    print("=" * 46)
    print(" EME Experiment Data Management System")
    print(f" http://localhost:{port}")
    print(" Close this window to stop the service.")
    print("=" * 46)
    print()

    threading.Thread(target=open_browser_when_ready, args=(port,), daemon=True).start()

    from streamlit.web import cli as stcli
    sys.argv = [
        "streamlit", "run", app_path,
        "--server.headless=true",
        f"--server.port={port}",
        "--browser.gatherUsageStats=false",
        "--server.fileWatcherType=none",  # frozen 下禁用文件监视，避免 watchdog 报错
    ]
    return stcli.main()


if __name__ == "__main__":
    sys.exit(main())
