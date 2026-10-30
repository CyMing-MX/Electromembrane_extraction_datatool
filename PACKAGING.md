# EME 打包指南（给自己看的操作手册）

> 目标：改完代码 → 一条命令 → 得到 `dist/EME_Data/`（EME_Data.exe + _internal + data）
> 当前打包方式：**PyInstaller 6.x onedir**（启动 2~5 秒；单文件版 onefile 启动要 30 秒+，已弃用）

---

## 一、日常打包（最常用）

**方法1：双击脚本**
```
双击 eme_db/build_exe.bat
```
完成后产物在 `eme_db/dist/EME_Data/`，把整个文件夹拷走用。

**方法2：命令行（与脚本等价）**
```bat
cd <项目目录>\eme_db
%LocalAppData%\Programs\Python\Python311\python.exe -m PyInstaller ^
  --noconfirm --name EME_Data --onedir --console ^
  --distpath dist --workpath build --specpath . ^
  --icon "<项目目录>\emedatatool.ico" ^
  --add-data "app.py;eme_db" --add-data "config.py;eme_db" --add-data "parser.py;eme_db" ^
  --add-data "calculator.py;eme_db" --add-data "models.py;eme_db" --add-data "excel_builder.py;eme_db" ^
  --add-data "dashboard.py;eme_db" --add-data "reference.py;eme_db" --add-data "naming.py;eme_db" ^
  --add-data "settings.py;eme_db" --add-data "current.py;eme_db" ^
  --collect-all streamlit --collect-all plotly --collect-all altair ^
  --collect-all pyarrow --collect-all charset_normalizer ^
  --collect-all rfc3987_syntax --collect-all jsonschema_specifications ^
  --hidden-import simpleeval ^
  --exclude-module PyQt5 --exclude-module PySide2 --exclude-module PySide6 ^
  --exclude-module tkinter --exclude-module IPython ^
  --exclude-module torch --exclude-module torchvision --exclude-module torchaudio ^
  --exclude-module scipy ^
  launcher.py
```
打包后把 `data` 模板拷进去（脚本已自动做）：
```bat
mkdir dist\EME_Data\data
xcopy /e /i /y ..\data\experiments dist\EME_Data\data\experiments
xcopy /e /i /y ..\data\photo dist\EME_Data\data\photo
xcopy /e /i /y ..\data\current dist\EME_Data\data\current
copy /y ..\data\analyte.xlsx dist\EME_Data\data\
copy /y ..\data\SLM.xlsx dist\EME_Data\data\
```

## 二、打包前检查清单

1. **代码能跑**：双击 `run.bat`，把要改的功能都过一遍。
2. **换图标**：替换 `<项目目录>\emedatatool.ico`（必须是 .ico，256×256 最佳），
   或改打包命令里 `--icon` 的路径。
3. **新依赖了第三方库**：在打包命令里补 `--collect-all 库名`。
   判断方法：如果打包后报 `ModuleNotFoundError` 或找不到某数据文件，就是收集遗漏。

## 三、打包后自检（强烈建议）

```bat
cd dist\EME_Data
EME_Data.exe selftest
```
显示 `SELFTEST OK: ...` 即完整；显示 `SELFTEST FAIL` 按提示补 collect 参数。
再双击 `EME_Data.exe` 走一遍主要页面（计算器/新增实验/详情/汇总）。

## 四、给别人的东西

整个 `EME_Data` 文件夹（exe + _internal + data 三个）。对方无需装 Python。

## 五、踩过的坑（为什么命令长这样）

| 坑 | 原因与解法 |
|---|---|
| 启动报 `server.port does not work when global.developmentMode is true` | frozen 下 Streamlit 自动进开发模式；launcher.py 已设 `STREAMLIT_GLOBAL_DEVELOPMENT_MODE=false` |
| 计算器页报 `No module named 'pyarrow.vendored.version'` | pyarrow 子模块收集不全 → `--collect-all pyarrow` |
| selftest 报 `rfc3987_syntax\...lark` / `jsonschema_specifications\schemas` 找不到 | altair/streamlit 的依赖数据文件 → 分别 `--collect-all` |
| 想排除大体积无用组件 | Qt 后端（matplotlib 用不到）→ launcher 设 `MPLBACKEND=Agg` + `--exclude-module PyQt5` 等（省约 170MB） |
| 产物突然大了 400MB（torch 进了 `_internal`） | matplotlib `cbook.py` 内部有函数级可选 `import torch`（张量判断），环境里装过 torch 后 PyInstaller 就会整体收集（371MB）；scipy 同理被可选分支扫入（~70MB）。运行时都不用 → `--exclude-module torch torchvision torchaudio scipy`（2026-10-30 已加入脚本；自检+实测无影响） |
| 单文件版（--onefile）启动 30 秒 | onefile 每次启动解压全部资源，无法优化 → 改用 6.x onedir（顶层只有 exe/_internal/data，不散乱且秒开） |
| 双击闪退报找不到 python311.dll | 多为杀毒误报：整个文件夹加白名单；用 selftest 判断文件是否完整 |
| `--add-data "源;目标"` 打包后 app.py 在哪 | 6.x onedir 在 `exe旁\_internal\eme_db\app.py`（launcher 用 `sys._MEIPASS` 定位，已兼容 5/6 两种布局） |
| 打包时报 dist 目录被占用 | 先关掉正在运行的 EME_Data.exe（任务管理器/`Get-Process EME_Data \| Stop-Process -Force`），必要时换个 `--distpath` |

## 六、版本记录

- 2026-10-30：脚本验证仍适用；新增排除 torch/scipy（环境装了 torch 后 matplotlib 可选导入把它拉进包，产物 +440MB）；Pillow 随图片功能自动收集，无需额外参数
- 2026-08-30：PyInstaller 6.22.2 onedir；启动 ~2s；裁剪 Qt 后端；图标 emedatatool.ico
- 历史：onefile（30s 启动，弃用）；PyInstaller 5.7 onedir（文件散落，弃用）
