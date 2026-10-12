# Excel SQL 查询台（按路径引用模式）
<img width="1900" height="909" alt="image" src="https://github.com/user-attachments/assets/0f3bcc1a-58b8-4c56-82ec-12152e1908d5" />

一个**本地运行**的 Excel / CSV / Doris 联合查询工具：静态前端 + FastAPI/DuckDB 后端，**按本地文件路径直接引用数据（不复制文件）**，自动展示字段，用 SQL 检索、关联、聚合数据，还能直连 **Apache Doris**，把「Excel 文件 + 数据库表」放在同一个查询里做跨源 JOIN。

## ✨ 特性

- **按路径直读，零复制**：Windows 原生「选择文件」对话框（tkinter）、内置文件浏览器或直接粘贴路径三种方式选文件；后端只读原文件、不产生副本，改了 Excel 不用重新选择，下次查询即为最新数据。
- **多格式源**：`.xlsx / .xls / .xlsb / .csv`；CSV 自动做 GBK→UTF-8 转码兜底，中文 Excel 导出不乱码。
- **单文件 / 多源关联**：单文件用表名 `data`；多源为每个文件起别名（`t1`、`t2`…），可跨文件 SQL JOIN。
- **Doris 直连**：走 MySQL 协议（默认 9030 端口）只读 ATTACH，一次连接暴露全部库/表；字段区展示列注释与键列标记；SQL 支持 `FROM 别名.库.表` 三段引用，Doris 表可与本地文件同台 JOIN。
- **Notebook 多 cell**：多条 SQL 各自运行/取消；per-cell 行数限制、折叠/删除、大纲浮轨导航、虚拟滚动结果表、底部合计行。
- **Parquet 查询缓存**：xlsx/xls/xlsb 首次读取后落盘为 parquet（`cache/` 目录），缓存按「路径+工作表+大小+修改时间」自动失效；选完文件即后台预构建，二次查询近毫秒级。
- **查询体验**：异步任务 + 真实进度百分比，可取消，300s 硬超时保护；结果默认 100 行、上限 10000 行（可逐 cell 调整）。
- **结果导出**：一键导出 `.xlsx / .csv`（原生「另存为」对话框；大结果集按 task_id 走后端缓存直读，不经前端回传）。
- **最近打开**：历史文件与多源「组合」一键恢复，路径失效批量校验（标红/自动移除），可回填上次 SQL。
- **本地持久化**：历史、Doris 连接、SQL 会话统一落 SQLite（`%LOCALAPPDATA%\ExcelSqlConsole\app.sqlite`），不依赖浏览器缓存。
- **编辑器细节**：全角/半角标点高亮（CodeMirror overlay），避免 SQL 逗号/括号混用不直观；`Ctrl+Enter` 运行、`Ctrl+/` 注释、`Ctrl+滚轮` 缩放。

> 为什么不用浏览器自带的 `<input type=file>`：浏览器出于安全设计，选择文件时
> **拿不到文件的本地绝对路径**（只有文件名），无法告知后端「原文件在哪」。
> 但由于**后端进程就运行在本机**，它可以直接调用 Windows 原生文件对话框
> （tkinter）弹出「选择文件」窗口，选完把真实路径返回给网页——效果等同系统文件管理器。
> 这是浏览器安全边界下的最佳方案。

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

# 4a. 浏览器模式（源码开发）
python backend/app.py
#     或：uvicorn app:app --app-dir backend

# 4b. 桌面窗口模式（pywebview 原生窗口）
python desktop.py
```

浏览器模式启动后访问：**http://127.0.0.1:8000**

> 源码运行时首次使用会联网自动 `INSTALL` DuckDB 扩展（读取 Excel 用 `excel`，连 Doris 用 `mysql`）；
> 打包产物则随包内置离线扩展（`duckdb_ext/`），完全离线可用。

## 使用说明

1. 点击 **「📂 选择文件」** → 弹出 Windows 原生文件对话框，选择 `.xlsx / .xls / .xlsb / .csv` 文件（也支持内置文件浏览器或直接粘贴路径）。
2. 选中后自动读取字段，左侧列出字段名与类型，点击字段可插入到 SQL 框。
3. 单文件模式下表名固定为 `data`：
   ```sql
   SELECT * FROM data LIMIT 10;
   SELECT 地区, SUM(销售额) AS 总销售额 FROM data GROUP BY 地区;
   ```
4. 多源模式：为每个源起别名，可跨源 JOIN（文件与 Doris 表混搭）：
   ```sql
   SELECT t1.客户, t2.销售额
   FROM t1 JOIN t2 ON t1.客户编码 = t2.客户编码;
   ```
5. Doris 源：在「连接数据库」里填 host / 端口（默认 9030）/ 账号即可，字段区下拉选库选表；SQL 里推荐写三段引用 `FROM 别名.库.表`（连接级源不建默认视图）。
6. 支持多条 SQL 各自成 cell；点击「运行」或按 `Ctrl+Enter`，结果以表格展示（可导出 xlsx / csv）。

## API 接口

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| POST | `/api/pick_file` | 弹出 Windows 原生「选择文件」对话框，返回真实路径（`{ path, cancelled }`） |
| POST | `/api/browse` | `{ "path": "D:\\..." }` 列出盘符/子目录/Excel 文件（不传 path 返回盘符列表） |
| POST | `/api/check_paths` | 批量校验路径是否仍存在（供「最近打开」标红失效条目） |
| POST | `/api/open` | `{ "path": ..., "sheet": null }` 返回字段列表、工作表列表与分阶段耗时；同时后台预构建缓存 |
| POST | `/api/describe` | 多源模式批量读取每个源的字段与工作表（含 Doris 源） |
| POST | `/api/db_probe` | 测试 Doris 连接并列出库/表（`{ ok, dbs, tables, server }`） |
| POST | `/api/query` | 提交查询任务（只读白名单校验 + 别名/表函数拦截），立即返回 `{ task_id }` |
| POST | `/api/progress` | 轮询任务进度：`{ status, phase, progress, result?, error? }`，done 时附带结果 |
| POST | `/api/cancel` | 取消运行中的查询（立即 interrupt 阻塞中的执行） |
| POST | `/api/export` | 弹出「另存为」对话框，把查询结果导出为 Excel(.xlsx) 或 CSV（优先按 task_id 直读后端缓存） |
| POST | `/api/prebuild_status` | 查询某文件的后台缓存预构建状态（`running / done / error / none`） |
| POST | `/api/cache/info` | 查询 parquet 缓存目录信息（数量/占用/位置） |
| POST | `/api/cache/clear` | 清空全部 parquet 缓存（下次查询自动重建） |
| POST | `/api/store/get` | 从本地 SQLite 批量读取持久化键（历史 / Doris 连接 / SQL 会话） |
| POST | `/api/store/set` | 批量写入/删除持久化键（value 为 null 即删除） |
| GET | `/api/health` | 健康检查 |

## 目录结构

```
excel-sql-web/
├── backend/
│   ├── app.py              # FastAPI + DuckDB：查询引擎、Doris 直连、Parquet 缓存、SQLite 持久化
│   └── __init__.py
├── frontend/               # 纯静态前端（无构建步骤）
│   ├── index.html
│   ├── style.css
│   ├── app.js
│   ├── codemirror.js / codemirror.css   # vendored CodeMirror 5
│   ├── mode/sql/           # SQL 语法高亮模式
│   └── addon/              # comment / hint 等插件
├── desktop.py              # 桌面启动器：单实例互斥锁 → 选空闲端口 → 后台 uvicorn → pywebview 窗口
├── build.py                # 打包脚本（PyInstaller onedir）
├── packaging.spec          # onedir 打包配置
├── make_icon.py            # 应用图标生成
├── assets/                 # 图标资源
├── duckdb_ext/             # 离线 DuckDB 扩展（excel / mysql_scanner），打包时随包携带（已 gitignore）
├── requirements.txt
├── ARCHITECTURE.md         # 架构详解（Mermaid 总览 + 组件职责 + 数据流）
└── README.md
```

运行期生成（已 gitignore）：`cache/`（Parquet 缓存，可随时手动清理）、`app.log`（后端日志）。

## 环境要求

> ⚠️ 不要用 **Python 3.8**：DuckDB 从 1.2.0 起不再提供 3.8 的预编译 wheel，
> 而本项目 `requirements.txt` 固定 `duckdb>=1.4.0`，在 3.8 上会回退源码编译，
> 通常报 `error: command ... link.exe ... failed: None`。

- **推荐**：Python **3.10 ~ 3.12**（3.9 也可以）。
- 确认只装 wheel、禁止编译源码：
  ```bash
  python --version
  pip install --only-binary :all: "duckdb>=1.4.0"
  # 若提示 "Could not find a version that satisfies..." 就说明该 Python 版本没有对应 wheel
  ```
- `excel` 扩展自 DuckDB **1.1.0** 引入；升级到 ≥1.4 修复了 Windows 中文控制台下首次下载扩展的 GBK 解码 bug。

## 打包为 Windows 桌面应用

源码开发用浏览器访问；也可打包成**双击即用的桌面应用**（内嵌后端，离线可用）。优先弹 `pywebview`（WebView2）原生窗口；若当前打包环境无法初始化窗口组件（pythonnet/.NET 桥在 PyInstaller 冻结产物中的已知限制），自动降级为**系统浏览器打开内嵌服务**，功能与原生窗口完全一致，见下文「浏览器降级」。

```bash
# 在 venv 中安装打包依赖（版本已实测钉住；pythonnet/clr_loader 版本必须匹配）
.venv\Scripts\python.exe -m pip install "pythonnet==3.0.3" "clr_loader==0.2.6" pywebview pyinstaller

# 一键打包（onedir 文件夹，产物在 dist\ExcelSqlConsole\）
.venv\Scripts\python.exe build.py
```

> ⚠️ 打包依赖组合务必按上面的版本钉住：pythonnet 3.0.x 与 clr_loader 0.2.x 强耦合；
> 已实测 3.1.0 在此环境完全不可用（venv 与冻结均报 `Failed to initialize Python.Runtime.dll`）。

流程：`desktop.py` 是桌面入口 —— 单实例互斥锁（重复双击只唤起已有实例）→ 选空闲端口后台启动 FastAPI → 待后端就绪后弹 `pywebview`（WebView2）窗口 → 关闭窗口时优雅停服务。

**浏览器降级（窗口不可用时的自动兜底）**：若窗口组件初始化失败（冻结产物的 pythonnet/.NET 桥初始化失败是已知环境限制，诊断输出里的 `webview/edgechromium: IMPORT FAILED` 即此情形），应用不会退出或报错，而是：
1. 自动用系统浏览器打开内嵌服务（`http://127.0.0.1:<端口>/`），页面功能与原生窗口完全一致；
2. 弹出说明对话框，「退出应用 / 保持运行」二选一，默认「保持运行」——服务继续在浏览器中工作；
3. 空闲 30 分钟无请求自动退出；使用中再次双击 exe 会把已有实例的页面重新在浏览器打开；
4. 失败细节只写 `%LOCALAPPDATA%\ExcelSqlConsole\app.log`（或 exe 同级 `app.log`），不打扰用户。

| 事项 | 处理 |
| --- | --- |
| 前端/扩展资源定位 | `backend/app.py` 的 `_resource_base()`：冻结后用 `sys._MEIPASS`，源码用项目目录 |
| 文件选择/另存对话框 | 进程内 `tkinter`（不用 `sys.executable -c` 子进程，冻结后会坏） |
| DuckDB 扩展 | 随包携带 `duckdb_ext/excel.duckdb_extension` 与 `mysql_scanner.duckdb_extension`，离线按绝对路径加载 |
| 端口 | 动态空闲端口（`find_free_port()`），不写死 8000；端口记 `%LOCALAPPDATA%\ExcelSqlConsole\last_port.txt`（或 exe 同级）供再次双击唤醒 |
| 单实例 | `Local\` 命名空间互斥锁（`Global\` 在受限令牌下会创建失败导致静默退出，已改） |
| 打包配置 | `packaging.spec`：`console=False` 窗口化，`collect_all("webview")` 收集 WebView2 .NET 程序集，`datas` 附带 `pythonnet/runtime` |
| 降内存 | monkey-patch WebView2 启动参数（`--disable-gpu`、`--renderer-process-limit=1`） |
| 后台日志 | 崩溃/启动失败写 `%LOCALAPPDATA%\ExcelSqlConsole\app.log` |

产物自检：`dist\ExcelSqlConsole\ExcelSqlConsole.exe --diagnose` 会卸载窗口、验证离线扩展与前端链路，结果写 `dist\ExcelSqlConsole\diagnose.txt`（或 `%LOCALAPPDATA%\ExcelSqlConsole\diagnose.txt`）；`--dllprobe` 可进一步定位 pythonnet 冻结加载失败的位点（复制到多目录逐一尝试），排查时用。

**代码签名与 SmartScreen**：PyInstaller 产物无 Authenticode 签名，双击会弹「Microsoft Defender SmartScreen 阻止了无法识别的应用启动」。运行一次签名脚本即可在本机消除该提示（发布者显示 `ExcelSql Console`）：

```powershell
# 生成自签名代码签名证书（首次，5 年有效）→ 签名（带时间戳）→ 加入本机受信任的根/发布者
powershell -ExecutionPolicy Bypass -File scripts\sign.ps1
```

> 脚本幂等：证书只生成一次，重复运行会复用；**每次重新打包后重跑一次**即可保证产物签名最新。
> 脚本会把签名证书加入"受信任的根证书颁发机构"，这一步 Windows 会弹安全确认框需点一次"是"；若运行后仍提示根证书未信任，双击脚本导出的 `ExcelSqlConsole_trust.cer` → 安装证书 → 选"受信任的根证书颁发机构" → 确定即可。

> 浏览器降级提示**只在首次出现**：选一次「保持运行」后记入 `%LOCALAPPDATA%\ExcelSqlConsole\fallback_silent.flag`，之后打包产物每次启动直接静默用浏览器打开，不再弹「窗口组件不可用」说明框。

> 说明：签名与信任均作用于**本机/当前用户**；产物拷到其他电脑仍会提示未知发布者，需要在那台机器执行同一脚本，或购买商业代码签名证书（OV/EV）实现全局可信。

> 依赖 **WebView2 Runtime**（Win10/11 及新版 Edge 自带；缺失时 pywebview 会自动下载安装）。
> onedir 文件夹需整体分发；若要单文件 exe 可改 spec 的 `exclude_binaries` 为打包内嵌（启动更慢、更易被杀软误报）。

## 注意事项

- **只读 SQL 白名单**：仅允许 `SELECT / WITH / DESCRIBE / SHOW / SUMMARIZE / PRAGMA`，避免误执行写操作。
- **禁止表函数**：用户 SQL 中不允许出现 `read_xlsx / read_parquet / glob` 等读文件函数，防止绕过路径白名单读取任意文件。
- **别名规则**：字母/数字/下划线，不以数字开头，且不能是 SQL 保留字。
- **行数限制**：单条查询默认返回 100 行、上限 10000 行（每个 cell 可单独调整），防止超大结果拖垮浏览器。
- **混合类型兜底**：Excel 某列「数字夹杂文本」被误推断类型时，后端自动整表按文本（all_varchar）重读，不丢数据；仍失败时按报错指引用 `CAST` 强制转换。
- **超时与保留**：查询硬超时 300 秒（可取消）；结果在服务内存保留 10 分钟供导出，过期需重新查询。
- **按路径引用**：文件被移动/删除后查询会报「文件不存在」，不会在 `backend/uploads/` 保存副本。
- **Doris 只读**：Doris 走 `ATTACH ... READ_ONLY`，只读不写远程库；首次连接会关闭 mysql 扩展默认事务包裹以兼容 Doris。

## 深入阅读

- [ARCHITECTURE.md](ARCHITECTURE.md) —— 整体架构总览（Mermaid 组件图、各模块职责、一次查询的完整数据流、Parquet 缓存设计）。