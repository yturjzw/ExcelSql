# Excel SQL 查询台（按路径引用模式）

一个静态前端 + Python 后端的本地查询工具：**按本地文件路径直接引用 Excel**（不复制文件），自动展示字段，直接用 SQL 检索数据。

- 后端：FastAPI + DuckDB（`excel` 扩展），通过 `read_xlsx` **直接读取原文件路径**，不做复制上传。
- 前端：纯 HTML/CSS/JS 静态页面，点击**「选择文件」弹出 Windows 原生文件对话框**（或内置文件浏览器点选），字段点选、Ctrl+Enter 运行。
- 用户在 SQL 中固定使用表名 **`data`** 引用打开的 Excel（默认第一个工作表）。

> 为什么不用浏览器自带的 `<input type=file>`：浏览器出于安全设计，选择文件时
> **拿不到文件的本地绝对路径**（只有文件名），无法告知后端「原文件在哪」。
> 但由于**后端进程就运行在本机**，它可以直接调用 Windows 原生文件对话框
> （tkinter）弹出「选择文件」窗口，选完把真实路径返回给网页——效果等同系统文件管理器。
> 这是浏览器安全边界下的最佳方案。

## 目录结构

```
excel-sql-web/
├── backend/
│   └── app.py          # FastAPI 服务 + DuckDB 按路径查询逻辑
├── frontend/
│   ├── index.html
│   ├── style.css
│   └── app.js
├── requirements.txt
└── README.md
```

## Python 版本要求

> ⚠️ 重要：不要用 **Python 3.8** 搭配最新版 DuckDB。DuckDB 从 1.2.0 起不再提供 3.8 的预编译 wheel，
> pip 会回退到源码编译，通常报 `error: command ... link.exe ... failed: None`。

- **推荐**：Python **3.10 ~ 3.12**（3.9 也可以）。
- **只能用 3.8 时**：`requirements.txt` 已把 duckdb 钉在 `>=1.1.0,<1.2.0`，`excel` 扩展仍可用。

确认是否走编译（只装 wheel、禁止编译源码）：

```bash
python --version
pip install --only-binary :all: "duckdb>=1.1.0,<1.2.0"
# 若提示 "Could not find a version that satisfies..." 就说明该 Python 版本没有对应 wheel
```

## 快速开始

```bash
# 1. 进入目录
cd excel-sql-web

# 2. （推荐）创建虚拟环境
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS / Linux:
source .venv/bin/activate

# 3. 安装依赖
pip install -r requirements.txt

# 4. 启动（首次会自动联网下载 excel 扩展）
python backend/app.py
# 或：
uvicorn app:app --app-dir backend
```

启动后浏览器访问：**http://127.0.0.1:8000**

## 使用说明

1. 点击左侧 **「📂 选择文件」** 按钮 → 弹出 Windows 原生文件对话框，选择 `.xlsx / .xls / .xlsb` 文件。
2. 选中后自动读取字段，左侧列出字段名和类型，点击字段可插入到 SQL 框。
3. 在右侧用 SQL 查询，表名固定为 `data`：
   - `SELECT * FROM data LIMIT 10;`
   - `SELECT 地区, SUM(销售额) AS 总销售额 FROM data GROUP BY 地区;`
4. 点击「运行查询」或按 `Ctrl+Enter`，结果以表格展示。

> 也支持内置文件浏览器：下方地址栏旁的「浏览」可逐级点进目录选文件；
> 或直接粘贴完整路径后点「读取字段」。文件被按地址直接读取，
> **修改原 Excel 后无需重新选择**，下次查询即为最新数据。

## API 接口

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| POST | `/api/pick_file` | 后端在本机弹出原生「选择文件」对话框，返回所选文件真实路径（`{ path, cancelled }`） |
| POST | `/api/browse` | `{ "path": "D:\\...\\dir" }`，列出盘符/子目录/Excel 文件（不传 path 则返回盘符列表） |
| POST | `/api/open` | `{ "path": "D:\\data\\sales.xlsx", "sheet": null }`，返回字段列表（不复制文件） |
| POST | `/api/query` | `{ "path": "D:\\data\\sales.xlsx", "sql": "SELECT ...", "sheet": null }`，返回 `{ columns, rows, row_count }` |
| GET | `/api/health` | 健康检查 |

## 注意事项

- 需要 **DuckDB ≥ 1.1.0**（`excel` 扩展在此版本加入）；**Python 3.8 必须用 `<1.2.0`**（见上文「Python 版本要求」）。`read_xlsx` 会先用采样推断类型；若某列推断出错，可在 SQL 里用 `CAST` 强制转换。
- 单次查询结果默认最多返回 **10000 行**，防止超大结果拖垮浏览器。
- 后端对 SQL 做了白名单限制（仅 `SELECT / WITH / DESCRIBE / SHOW / SUMMARIZE / PRAGMA`），避免误执行写操作。
- **按路径引用不复制文件**：后端只读取原文件，不会在 `backend/uploads/` 保存副本。若文件被移动/删除，查询会报「文件不存在」。

## 扩展思路

- 支持多工作表：后端 `read_xlsx` 已支持 `sheet` 参数，前端输入框可填工作表名。
- 结果导出：把查询结果 `COPY ... TO 'out.csv'` 或转成 Parquet。
- 多个文件联表：`read_xlsx` 支持通配符路径（如 `D:\data\2024-*.xlsx`）。

## 打包为 Windows 桌面应用

源码开发时用浏览器访问；也可打包成**双击即用的原生桌面窗口**（内嵌后端，无浏览器地址栏，离线可用）。

```bash
# 在 venv 中安装打包依赖
.venv\Scripts\python.exe -m pip install pywebview pyinstaller

# 一键打包（onedir 文件夹，产物在 dist\ExcelSqlConsole\）
.venv\Scripts\python.exe build.py
```

流程：`desktop.py` 是桌面入口 —— 单实例互斥锁（重复双击只唤起已有窗口）→ 选空闲端口后台启动 FastAPI → 待后端就绪后弹 `pywebview`（WebView2）窗口 → 关闭窗口时优雅停服务。

关键实现点（打包环境与源码环境的差异）：

| 事项 | 处理 |
| --- | --- |
| 前端/扩展资源定位 | `backend/app.py` 的 `_resource_base()`：冻结后用 `sys._MEIPASS`，源码用项目目录 |
| 文件选择对话框 | 进程内 `tkinter`（不再用 `sys.executable -c` 子进程，冻结后会坏） |
| DuckDB `excel` 扩展 | 随包携带 `duckdb_ext/excel.duckdb_extension`，`_load_excel()` 按绝对路径离线加载 |
| 端口 | 动态空闲端口（`find_free_port()`），不再写死 8000 |
| 打包产物 | `packaging.spec`：`console=False` 窗口化，`collect_all("webview")` 收集 WebView2 .NET 程序集 |
| 后台日志 | 崩溃/启动失败写 `%LOCALAPPDATA%\ExcelSqlConsole\app.log` |

产物自检：`dist\ExcelSqlConsole\ExcelSqlConsole.exe --diagnose` 会卸载窗口、验证离线扩展与前端链路，结果写 `dist\ExcelSqlConsole\diagnose.txt`（或 `%LOCALAPPDATA%\ExcelSqlConsole\diagnose.txt`）。

> 依赖 **WebView2 Runtime**（Win10/11 及新版 Edge 自带；缺失时 pywebview 会自动下载安装）。
> onedir 文件夹需整体分发；若要单文件 exe 可把 spec 的 `exclude_binaries` 改为打包内嵌（启动更慢、更易被杀软误报）。