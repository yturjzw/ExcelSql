"""Excel SQL Query 后端（按路径引用模式）。

提供接口：
  POST /api/open    —— 传入 Excel 本地路径，返回字段名列表（按地址引用，不复制文件）
  POST /api/query   —— 按路径引用 Excel 执行 SQL，返回 {columns, rows, row_count}
  GET  /api/health  —— 健康检查

基于 FastAPI + DuckDB 的 excel 扩展。上传的文件不会被复制保存，
DuckDB 通过 read_xlsx 直接读取原文件路径（按地址引用）。
"""
from __future__ import annotations

import asyncio
import hashlib
import html
import json
import os
import re
import socket
import sqlite3
import string
import sys
import tempfile
import threading
import time
import tkinter as tk
import uuid
import zipfile
from pathlib import Path
from tkinter import filedialog

import duckdb
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

def _resource_base() -> Path:
    """资源根目录：源码运行时为项目根；PyInstaller 冻结后为 sys._MEIPASS。"""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent.parent


BASE_DIR = _resource_base()

# 前端静态文件目录
FRONTEND_DIR = BASE_DIR / "frontend"

# 打包时嵌入的 DuckDB 扩展目录（内含 v{version}/windows_amd64/excel.duckdb_extension）
# 用于离线 LOAD excel / mysql_scanner，摆脱对用户机器 .duckdb 缓存的依赖
BUNDLED_EXT_DIR = BASE_DIR / "duckdb_ext"

# 支持的 Excel 扩展名
ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".xlsb", ".csv"}

# 单次查询默认返回行数（用户可在前端配置）
DEFAULT_LIMIT = 100

# 单次查询最多返回行数，防止超大结果集拖垮浏览器
MAX_RESULT_ROWS = 10000

# 用户在 SQL 中引用的视图名（指向按路径引用的 Excel 的第一个工作表）
TABLE_ALIAS = "data"

app = FastAPI(title="Excel SQL Query (path reference)")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# ---------- SQLite 持久化（前端 localStorage 迁移目标） ----------
# 本地桌面应用：历史文件、多源组合、Doris 连接、SQL 会话统一落 SQLite 数据库，
# 摆脱 WebView2 profile/localStorage 的绑定（清缓存/换 profile 不丢、可备份、可查询）。
# 库文件放在 %LOCALAPPDATA%\ExcelSqlConsole\app.sqlite（与 app.log / diagnose.txt 同目录）。
_STORE_DB: "sqlite3.Connection | None" = None
_STORE_LOCK = threading.Lock()


def _store_dir() -> Path:
    """SQLite 库文件目录：优先 LOCALAPPDATA；不可用则回退 exe 所在目录（便于排查）。"""
    local = os.environ.get("LOCALAPPDATA")
    if local:
        p = Path(local) / "ExcelSqlConsole"
        try:
            p.mkdir(parents=True, exist_ok=True)
            return p
        except OSError:
            pass
    return Path(os.path.dirname(sys.executable)) if getattr(sys, "frozen", False) else Path(__file__).resolve().parent.parent


def _get_store() -> "sqlite3.Connection":
    """返回进程级 SQLite 连接（首次打开建表）；后续读写复用。"""
    global _STORE_DB
    if _STORE_DB is None:
        db_path = _store_dir() / "app.sqlite"
        _STORE_DB = sqlite3.connect(str(db_path), check_same_thread=False)
        _STORE_DB.execute("PRAGMA journal_mode = WAL")
        _STORE_DB.execute(
            "CREATE TABLE IF NOT EXISTS kv (k TEXT PRIMARY KEY, v TEXT NOT NULL)"
        )
        _STORE_DB.commit()
    return _STORE_DB


def _store_get(key: str) -> str | None:
    """读取一个 key 的 JSON 字符串；不存在返回 None。"""
    with _STORE_LOCK:
        con = _get_store()
        row = con.execute("SELECT v FROM kv WHERE k = ?", (key,)).fetchone()
    return row[0] if row else None


def _store_set(key: str, value: str) -> None:
    """写入一个 key 的 JSON 字符串（upsert）。"""
    with _STORE_LOCK:
        con = _get_store()
        con.execute(
            "INSERT INTO kv (k, v) VALUES (?, ?) "
            "ON CONFLICT(k) DO UPDATE SET v = excluded.v",
            (key, value),
        )
        con.commit()


def _store_remove(key: str) -> None:
    """删除一个 key。"""
    with _STORE_LOCK:
        con = _get_store()
        con.execute("DELETE FROM kv WHERE k = ?", (key,))
        con.commit()


def _close_store() -> None:
    """应用退出时关闭 SQLite 连接（desktop.shutdown() → app.shutdown() 调用）。"""
    global _STORE_DB
    if _STORE_DB is not None:
        try:
            _STORE_DB.close()
        except Exception:
            pass
        _STORE_DB = None


# ---------- 工具函数 ----------

def _reader_name(file_path: Path) -> str:
    """根据扩展名返回 DuckDB excel 扩展的读取函数。"""
    ext = file_path.suffix.lower()
    if ext == ".xlsx":
        return "read_xlsx"
    if ext == ".xls":
        return "read_xls"
    if ext == ".csv":
        return "read_csv"
    return "read_xlsb"


def _is_cell_conversion_error(e: Exception) -> bool:
    """判断是否为 excel 扩展的「单元格文本无法转换为推断类型」错误。

    DuckDB 的 excel 扩展按采样推断列类型：当一列大多数是数字、个别单元格是文本
    （典型如编码 KE07）时，会把整列推断为 DOUBLE，并在扫描到文本单元格时抛出
    「read_xlsx: Failed to parse cell ...: Could not convert string ... to DOUBLE」。
    这类错误不是用户 SQL 的问题，可通过把该文件重读为全文本（all_varchar）解决。

    只认「读取阶段」的特征（Failed to parse cell，或 read_xlsx 伴随 could not
    convert），避免误把用户 SQL 里显式 CAST 的转换错误也判为此类而触发重读。
    """
    s = str(e).lower()
    return "failed to parse cell" in s or ("could not convert" in s and "read_xlsx" in s)


# CSV 编码检测顺序：UTF-8（含 BOM）优先，失败则按 GBK（中文 Excel 导出默认编码）
_CSV_ENCODINGS = ("utf-8-sig", "gbk")

# 头部采样字节数：用于编码探测。表头 + 前几十行数据通常都在此范围内，
# 足够判断编码；超出头部之后才出现的中文在真实 CSV 里几乎不存在。
_CSV_HEAD_BYTES = 262144  # 256 KB


def _prepare_csv(file_path: Path) -> Path:
    """检测 CSV 编码；非 UTF-8 时转码为临时 UTF-8 文件再交给 DuckDB。

    DuckDB 的 read_csv 只支持 UTF-8，而中文 Windows 的 Excel「另存为 CSV」
    默认是 GBK 编码，直接读会报 Invalid unicode。这里用 Python 做兜底转码。

    P0 优化：只读文件头部（256KB）做 UTF-8 探测。纯 UTF-8/ASCII 的大文件
    无需整体读入内存再解码——原实现 read_bytes() 会让 890MB 级 CSV 在「编码
    检测」这一步就白白全量读一遍，拖慢首次查询。仅当头部判为「非 UTF-8」
    （典型 GBK）时才全量读入并转码。
    """
    with file_path.open("rb") as f:
        head = f.read(_CSV_HEAD_BYTES)

    # 截断安全：多字节字符可能恰好跨 head 末尾被切断，去掉末尾 0~3 字节逐次重试
    # （UTF-8 单字符最长 4 字节）。小文件时 head 即全文件，无截断问题。
    is_utf8 = False
    for cut in range(4):
        try:
            head[: len(head) - cut].decode("utf-8")
            is_utf8 = True
            break
        except UnicodeDecodeError:
            continue
    if is_utf8:
        return file_path   # 合法 UTF-8 / ASCII（含 BOM），直接用原文件，不转码

    # 非 UTF-8：全量读入，按 utf-8-sig → gbk 顺序尝试转码
    data = file_path.read_bytes()
    last_err: UnicodeDecodeError | None = None
    for enc in _CSV_ENCODINGS:
        try:
            text = data.decode(enc)
            break
        except UnicodeDecodeError as e:
            last_err = e
    else:
        raise HTTPException(
            status_code=400,
            detail=f"CSV 编码无法识别（既不是 UTF-8 也不是 GBK）: {file_path.name}",
        ) from last_err

    # 转码后的临时文件（NamedTemporaryFile 关闭后可被 DuckDB 重复打开）。
    # 用二进制写，避免 Windows 文本模式把 \n 二次翻译成 \r\n（原文件已是 \r\n，
    # 会变成 \r\r\n 多出空行，导致 DuckDB 嗅探分隔符失败）。
    fd, tmp_name = tempfile.mkstemp(suffix=".csv", prefix="excel_sql_csv_")
    os.close(fd)
    tmp_path = Path(tmp_name)
    tmp_path.write_bytes(text.encode("utf-8"))
    return tmp_path


def _cleanup_tmp_files(paths: list[Path] | None) -> None:
    """删除 CSV 转码产生的临时文件（查询/探测结束后调用，失败静默）。"""
    if not paths:
        return
    for p in paths:
        try:
            p.unlink(missing_ok=True)
        except OSError:
            pass


def _load_excel(con: duckdb.DuckDBPyConnection) -> None:
    """加载 excel 扩展。

    打包后优先使用随包携带的扩展（离线可用、不依赖用户机器的 .duckdb 缓存）；
    源码运行时回退到默认缓存路径，首次缺失再 INSTALL 联网下载。
    注意：DuckDB 1.1.x 在 Windows 下首次下载扩展有编码 bug，建议升级到 >=1.4。
    """
    # 1) 随包扩展文件存在时，按绝对路径直接加载（离线、无需版本目录布局）
    bundled = BUNDLED_EXT_DIR / "excel.duckdb_extension"
    if bundled.is_file():
        try:
            con.execute(f"LOAD '{_sql_path(bundled)}'")
            return
        except Exception:
            pass  # 回退到默认缓存/联网安装

    # 2) 源码/开发环境：先 LOAD 命中本地缓存，失败则 INSTALL 联网下载
    try:
        con.execute("LOAD excel;")
    except Exception:
        con.execute("INSTALL excel;")
        con.execute("LOAD excel;")


def _load_mysql(con: duckdb.DuckDBPyConnection) -> None:
    """加载 mysql_scanner 扩展（Doris 走 MySQL 协议，靠它 ATTACH 直连）。

    与 _load_excel 同款策略：打包后优先随包离线 LOAD（mysql_scanner.duckdb_extension
    已复制进 duckdb_ext/）；源码运行时回退 LOAD，缺失则 INSTALL 联网下载。
    注：Doris 不认 DuckDB mysql 扩展默认的事务包裹（START TRANSACTION READ ONLY），
    必须关掉事务隔离，否则任何查询都会报 mismatched input 'READ'。
    """
    bundled = BUNDLED_EXT_DIR / "mysql_scanner.duckdb_extension"
    if bundled.is_file():
        try:
            con.execute(f"LOAD '{_sql_path(bundled)}'")
        except Exception:
            pass  # 回退到默认缓存/联网安装
    try:
        con.execute("LOAD mysql;")
    except Exception:
        con.execute("INSTALL mysql;")
        con.execute("LOAD mysql;")
    try:
        con.execute("SET mysql_enable_transactions = false")   # P0：Doris 兼容必须
    except Exception:
        pass


# ---------- 进程级共享连接（Doris ATTACH 复用） ----------
# 首次查询某 Doris 时建共享 DuckDB 连接并 ATTACH（~1.8s 一次性开销），之后查询
# 直接复用；含 Doris 源的任务持 _SHARED_LOCK 串行执行（共享连接上的视图/状态不能并发）。
# 应用退出（shutdown()）时统一关闭。
_SHARED_LOCK = threading.RLock()
_SHARED_CON: "duckdb.DuckDBPyConnection | None" = None
_DORIS_ATTACHED: dict[tuple, str] = {}   # conn_key -> attach_alias（已 ATTACH，进程级缓存）


def _get_shared_con() -> "duckdb.DuckDBPyConnection":
    """返回进程级共享 DuckDB 连接（首次创建：LOAD excel+mysql 扩展、SET 进度条）。"""
    global _SHARED_CON
    if _SHARED_CON is None:
        con = duckdb.connect()
        # 与 _run_query_task 原来对每个任务连接做的初始化一致，共享连接只做一次
        con.execute("SET enable_progress_bar = true")
        con.execute("SET enable_progress_bar_print = false")
        con.execute("SET progress_bar_time = 0")
        _load_excel(con)
        _load_mysql(con)   # 内部已 LOAD mysql_scanner + SET 关事务
        _SHARED_CON = con
    return _SHARED_CON


def _invalidate_shared_con() -> None:
    """丢弃共享连接（如 Doris 掉线需要重建）。必须在持有 _SHARED_LOCK 时调用。"""
    global _SHARED_CON
    if _SHARED_CON is not None:
        try:
            _SHARED_CON.close()
        except Exception:
            pass
        _SHARED_CON = None
    _DORIS_ATTACHED.clear()


def _close_shared_con() -> None:
    """应用退出：关闭共享连接（desktop.py 的 shutdown() 调用）。"""
    with _SHARED_LOCK:
        _invalidate_shared_con()


def _list_drives() -> list[str]:
    """列出 Windows 上可用的盘符（按字母顺序 A-Z 检测）。"""
    drives = []
    for letter in string.ascii_uppercase:
        if Path(f"{letter}:\\").exists():
            drives.append(f"{letter}:\\")
    return drives


def _list_directory(path_str: str) -> dict:
    """列出某个目录下的子目录和 Excel 文件（仅展示可点选的项）。"""
    if not path_str:
        # 未提供路径：返回盘符列表
        drives = _list_drives()
        items = [{"name": d, "path": d, "type": "drive"} for d in drives]
        return {"path": "", "parent": None, "items": items}

    raw = path_str.strip()
    # 兼容用户直接填盘符/目录
    p = Path(raw)
    if not p.exists():
        raise HTTPException(status_code=404, detail=f"路径不存在: {raw}")
    if not p.is_dir():
        raise HTTPException(status_code=400, detail=f"不是目录: {raw}")

    dirs = []
    files = []
    try:
        for entry in sorted(p.iterdir(), key=lambda x: x.name.lower()):
            try:
                if entry.is_dir():
                    dirs.append(entry)
                elif entry.suffix.lower() in ALLOWED_EXTENSIONS:
                    files.append(entry)
            except OSError:
                # 无权限或无响应的条目直接跳过
                continue
    except PermissionError as e:
        raise HTTPException(status_code=403, detail=f"没有权限访问该目录: {raw}") from e
    except OSError as e:
        raise HTTPException(status_code=500, detail=f"读取目录失败: {raw} -> {e}") from e

    items = []
    for d in dirs:
        items.append({"name": d.name, "path": str(d), "type": "dir"})
    for f in files:
        items.append({"name": f.name, "path": str(f), "type": "file"})

    # 构造父目录（根目录/盘符时 parent 为 None）
    parent = str(p.parent) if p.parent != p else None
    return {"path": str(p), "parent": parent, "items": items}


def _pick_file_via_tkinter() -> str | None:
    """弹出 Windows 原生「选择文件」对话框，返回选中的文件路径。

    后端进程就运行在本机，因此可以调用系统级对话框拿到真实路径（浏览器做不到）。
    tkinter 必须运行在创建它的线程中，且 askopenfilename 会阻塞到用户选择/取消；
    打包进 exe 后不能再用「sys.executable -c 子进程」模式（冻结环境下会失败），
    故改为进程内直接调用，由调用方（async endpoint）通过 asyncio.to_thread 放到
    线程池执行，避免阻塞 FastAPI 事件循环。
    """
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        path = filedialog.askopenfilename(
            title="选择数据文件",
            filetypes=[
                ("Excel/CSV 文件", "*.xlsx *.xls *.xlsb *.csv"),
                ("所有文件", "*.*"),
            ],
        )
    finally:
        root.destroy()
    return path or None


def _pick_save_path(default_name: str, filetypes: list[tuple[str, str]]) -> str | None:
    """弹出 Windows 原生「另存为」对话框，返回用户选择的保存路径（可取消）。"""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        path = filedialog.asksaveasfilename(
            title="导出查询结果",
            defaultextension=filetypes[0][1].split(".")[-1],
            initialfile=default_name,
            filetypes=filetypes,
        )
    finally:
        root.destroy()
    return path or None


def _resolve_path(path_str: str) -> Path:
    """校验并解析本地 Excel 路径（按地址引用）。"""
    # 允许用户粘贴时带引号，去掉两端引号与首尾空白
    raw = (path_str or "").strip().strip('"').strip("'").strip()
    if not raw:
        raise HTTPException(status_code=400, detail="路径不能为空")

    p = Path(raw).expanduser()
    if p.suffix.lower() not in ALLOWED_EXTENSIONS:
        raise HTTPException(status_code=400, detail="只支持 .xlsx / .xls / .xlsb / .csv 文件")
    if not p.is_file():
        raise HTTPException(status_code=404, detail=f"文件不存在或不是文件: {raw}")

    return p.resolve()


def _sql_path(file_path: Path) -> str:
    """把路径转成 SQL 字符串字面量内容：转正斜杠并转义单引号。"""
    return file_path.as_posix().replace("'", "''")


def _sheet_literal(sheet: str | None) -> str:
    """把 sheet 参数转成 SQL 字面量；空值表示读取第一个工作表。"""
    sheet = (sheet or "").strip()
    if not sheet:
        return ""  # 不传 sheet，read_xlsx 默认读第一个 sheet
    if sheet.isdigit():
        return f"{int(sheet)}"
    if re.search(r"['\"\\;\0]", sheet):
        raise HTTPException(status_code=400, detail="工作表名包含非法字符")
    return f"'{sheet}'"


def _list_sheets(file_path: Path) -> list[str]:
    """列出 .xlsx 工作簿的所有工作表名（按 workbook.xml 中的顺序）。

    通过标准库 zipfile 直接解析 xlsx（本质是 zip）里的 workbook.xml，
    不依赖 DuckDB 或其他第三方库；解析失败或非 xlsx 时返回空列表。
    """
    if file_path.suffix.lower() != ".xlsx":
        return []
    try:
        with zipfile.ZipFile(file_path) as zf:
            for name in ("xl/workbook.xml", "xl/Workbook.xml"):
                try:
                    xml = zf.read(name).decode("utf-8", "replace")
                    break
                except KeyError:
                    continue
            else:
                return []
    except (zipfile.BadZipFile, OSError):
        return []

    # 依次解析 <sheet ... name="xxx"/>，保留文档中的出现顺序
    sheets: list[str] = []
    for m in re.finditer(r'<sheet\b[^>]*?\bname="([^"]*)"', xml, re.IGNORECASE):
        sheets.append(html.unescape(m.group(1)))
    return sheets


# SQL 保留字黑名单：别名不能撞上（否则 SQL 里无法裸写引用）
_SQL_RESERVED = {
    "select", "from", "where", "join", "inner", "outer", "left", "right", "full",
    "cross", "on", "using", "group", "by", "order", "having", "limit", "offset",
    "union", "all", "distinct", "as", "and", "or", "not", "in", "is", "null",
    "like", "between", "exists", "case", "when", "then", "else", "end", "cast",
    "with", "view", "table", "create", "replace", "insert", "update", "delete",
    "drop", "into", "values", "set", "primary", "key", "foreign", "references",
    "index", "unique", "check", "default", "constraint", "asc", "desc", "true",
    "false", "if", "over", "partition", "window", "fetch", "first", "next", "rows",
    "only", "using", "recursive", "materialized",
}

# 用户 SQL 里禁止出现的表函数（否则可用 read_xlsx 绕过路径白名单直接读任意文件）
_FORBIDDEN_TABLE_FUNCS = {"read_xlsx", "read_xls", "read_xlsb", "read_csv",
                          "read_parquet", "read_json", "read_ndjson", "glob"}

_ALIAS_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]{0,63}$")


def _validate_alias(alias: str) -> str:
    """校验并规范化 SQL 别名：仅允许字母/数字/下划线、不以数字开头、非保留字。"""
    a = (alias or "").strip().strip('"')
    if not a:
        raise HTTPException(status_code=400, detail="数据源别名不能为空")
    if not _ALIAS_RE.match(a):
        raise HTTPException(status_code=400, detail=f"别名不合法: {alias}（仅允许字母/数字/下划线，且不以数字开头）")
    if a.lower() in _SQL_RESERVED:
        raise HTTPException(status_code=400, detail=f"别名不能是 SQL 保留字: {alias}")
    return a


def _assert_no_forbidden_tables(sql: str) -> None:
    """拦截用户 SQL 中直接调用表函数（读文件）绕过路径白名单的行为。"""
    lowered = sql.lower()
    for fn in _FORBIDDEN_TABLE_FUNCS:
        # 用词边界匹配，避免误伤普通标识符（如列名叫 read_csv_val 不会命中）
        if re.search(rf"\b{re.escape(fn)}\s*\(", lowered):
            raise HTTPException(status_code=400, detail=f"不支持的函数: {fn}()")


def _leading_sql_keyword(sql: str) -> str:
    """去掉行首空白与 SQL 注释后，返回第一条语句的首个关键字（小写）。

    允许 SQL 以 -- 行注释 / /* */ 块注释开头（或注释后紧跟 SELECT 等），
    避免「仅允许只读查询」的校验被各行注释误判。
    """
    s = sql
    while True:
        s = s.lstrip()
        if s.startswith("--"):
            nl = s.find("\n")
            s = s[nl + 1:] if nl != -1 else ""
        elif s.startswith("/*"):
            end = s.find("*/", 2)
            if end == -1:
                return ""  # 未闭合块注释：无有效关键字
            s = s[end + 2:]
        else:
            break
    if not s:
        return ""
    return s.lstrip().split(None, 1)[0].lower()


def _build_source_expr(file_path: Path, sheet: str | None = None, all_varchar: bool = False,
                       tmp_out: list[Path] | None = None) -> str:
    """构造 DuckDB「读取源」函数表达式（含 CSV 转码与 xlsx 全文本容错）。

    返回形如 read_xlsx('C:/data.xlsx', header=true, sheet='Sheet1', all_varchar=true)
    的 FROM 表达式。CSV 会先做 GBK→UTF-8 兜底转码。
    当 CSV 发生转码、产生临时文件时，把临时文件路径追加到 tmp_out（调用方负责
    在查询/探测完成后清理；DuckDB 惰性读，因此清理必须发生在 fetchall 之后）。
    """
    reader = _reader_name(file_path)
    if reader == "read_csv":
        prepared = _prepare_csv(file_path)
        if tmp_out is not None and prepared != file_path:
            tmp_out.append(prepared)
        file_path = prepared
        sheet_arg = ""
    else:
        sheet_part = _sheet_literal(sheet)
        sheet_arg = f", sheet={sheet_part}" if sheet_part else ""
    # xlsx 单元格文本被误推断为 DOUBLE 时，可整表按文本重读（见 _is_cell_conversion_error）
    all_text_arg = ", all_varchar=true" if (reader == "read_xlsx" and all_varchar) else ""
    return f"{reader}('{_sql_path(file_path)}', header=true{sheet_arg}{all_text_arg})"


# ---------- Parquet 查询缓存 ----------
# xlsx/xls/xlsb 是「zip 压缩的 XML」，每次 read_xlsx 都要完整解压+解析再扫描，
# 30 万行的真实表常需数秒～十几秒。首次读取时把解析结果落成 DuckDB 原生 Parquet
# （自带行数元数据、列式压缩、扫描极快），后续查询直接读 Parquet：
#   count 接近毫秒级，普通聚合/过滤也快一个数量级。
# 缓存键 = 文件路径 + sheet + 大小 + mtime（stat 极快），文件修改后自动失效重建。
# 缓存目录 = exe/项目根 同级的 cache/ 目录（不占用 C 盘系统目录）。

_CACHE_VERSION = "v1"


def _writable_base_dir() -> Path:
    """可写根目录：优先 exe（冻结）或项目根（源码）所在目录。

    缓存放在这里——跟可执行文件/项目在一起，不占用 C 盘系统目录。
    注意不能用 BASE_DIR（冻结时指向 sys._MEIPASS 临时解包目录，会被系统回收）。
    """
    if getattr(sys, "frozen", False) and sys.executable:
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def _cache_root() -> Path:
    """返回 parquet 缓存目录：可执行文件/项目根 同级的 cache/ 目录。

    - 打包后：<dist>\\ExcelSqlConsole\\cache\\（跟 exe 同级）
    - 源码运行：<项目根>\\cache\\
    目录不可写时（如 exe 放在只读位置），回退到系统临时目录。
    """
    try:
        root = _writable_base_dir() / "cache"
        root.mkdir(parents=True, exist_ok=True)
        # mkdir 成功不代表真正可写，做一次写入探测确保可用
        probe = root / ".write_probe"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink(missing_ok=True)
        return root
    except OSError:
        root = Path(tempfile.gettempdir()) / "ExcelSqlConsole_cache"
        try:
            root.mkdir(parents=True, exist_ok=True)
        except OSError:
            pass
        return root


def _source_cache_signature(resolved: Path, sheet: str | None, all_varchar: bool = False) -> str | None:
    """基于路径 + sheet + 读取语义(all_varchar) + 大小 + mtime 计算缓存签名。

    必须把 all_varchar 计入签名：正常类型读与全文本兜底读是两种不同语义，
    若共用同一缓存文件，会互相污染（见 _register_sources 的回退逻辑）。
    """
    try:
        st = resolved.stat()
    except OSError:
        return None
    raw = f"{_CACHE_VERSION}|{resolved.as_posix().lower()}|{sheet or ''}|{int(all_varchar)}|{st.st_size}|{st.st_mtime_ns}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def _cached_parquet_path(sig: str) -> Path:
    """按签名定位缓存文件（签名已含路径+sheet+语义+大小+mtime，唯一标识一份数据）。

    文件名不含别名：多源模式下别名是 t1/t2 等运行时动态值，若按别名命名，
    预构建（发生在 open，不知道未来别名）与查询（已确定别名）会使用不同文件名，
    导致缓存永远命中不了。去掉别名后，指向同一文件+sheet 的多个源还能共享缓存。
    """
    return _cache_root() / f"{sig}.parquet"


def _cached_parquet_exists(resolved: Path, sheet: str | None, alias: str, all_varchar: bool) -> bool:
    """该数据源是否已有对应语义的 parquet 缓存文件（仅 stat，不扫描源表）。"""
    if resolved.suffix.lower() == ".csv":
        return False
    sig = _source_cache_signature(resolved, sheet, all_varchar)
    if sig is None:
        return False
    return _cached_parquet_path(sig).is_file()


def _parquet_cache_load(con: duckdb.DuckDBPyConnection, resolved: Path,
                        sheet: str | None, alias: str, all_varchar: bool) -> tuple[Path | None, bool]:
    """返回 (parquet 路径 or None, 本次是否新建了缓存)。

    built=True 表示这次扫描源表并写入了新 parquet；False 表示命中已有缓存。
    CSV 不加缓存（本就很快），返回 (None, False)。失败抛异常（交上层回退）。
    """
    if resolved.suffix.lower() == ".csv":
        return None, False  # CSV 本来就快，跳过缓存

    sig = _source_cache_signature(resolved, sheet, all_varchar)
    if sig is None:
        return None, False
    parquet = _cached_parquet_path(sig)
    if parquet.is_file():
        return parquet, False

    # 未命中：把源表 COPY 成 parquet（这一次扫描即构建缓存）
    tmp = parquet.with_name(f"{parquet.name}.{uuid.uuid4().hex[:8]}.tmp")
    expr = _build_source_expr(resolved, sheet, all_varchar)
    try:
        con.execute(f"COPY (SELECT * FROM {expr}) TO '{_sql_path(tmp)}' (FORMAT PARQUET)")
        os.replace(str(tmp), str(parquet))
        return parquet, True
    except Exception:
        # 构建失败（如首次扫描命中单元格转换错误，交给 all_varchar 回退）——清理临时文件
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass
        raise


def _prebuild_key(resolved: Path, sheet: str | None) -> str:
    """预构建状态的键（与缓存签名不同，不含 mtime，便于跨文件稳定追踪）。"""
    return f"{resolved.as_posix().lower()}|{sheet or ''}"


# 后台预构建状态登记：key -> {"status": "running"|"done"|"error"}
_PREBUILD_STATUS: dict[str, dict] = {}
_PREBUILD_LOCK = threading.Lock()


def _prebuild_cache(resolved: Path, sheet: str | None, alias: str) -> None:
    """后台预构建缓存：选完文件后立刻在空闲期把首读+建缓存提前完成。

    独立 DuckDB 连接，尽力而为（成功则后续查询直接命中；失败静默，等真正
    查询时再按老路径兜底）。混合类型文件在普通类型扫描下会失败，这里同样
    先试普通类型、失败再试全文本，把两种语义的缓存都尽量备好。
    """
    if resolved.suffix.lower() == ".csv":
        return
    key = _prebuild_key(resolved, sheet)
    with _PREBUILD_LOCK:
        _PREBUILD_STATUS[key] = {"status": "running", "started": time.perf_counter()}
    try:
        # 已存在全文本缓存说明此前已判定为混合类型，无需再试普通类型（会白扫一遍）
        if not _cached_parquet_exists(resolved, sheet, alias, all_varchar=True):
            con = None
            try:
                con = duckdb.connect()
                _load_excel(con)
                try:
                    _parquet_cache_load(con, resolved, sheet, alias, all_varchar=False)
                except Exception:
                    # 普通类型失败（混合类型）→ 试全文本
                    try:
                        _parquet_cache_load(con, resolved, sheet, alias, all_varchar=True)
                    except Exception:
                        pass
            except Exception:
                pass
            finally:
                if con is not None:
                    try:
                        con.close()
                    except Exception:
                        pass
        with _PREBUILD_LOCK:
            st = _PREBUILD_STATUS.get(key) or {}
            st["status"] = "done"
            st["elapsed_ms"] = round((time.perf_counter() - st.get("started", time.perf_counter())) * 1000.0, 1)
    except Exception:
        with _PREBUILD_LOCK:
            st = _PREBUILD_STATUS.get(key) or {}
            st["status"] = "error"
            st["elapsed_ms"] = round((time.perf_counter() - st.get("started", time.perf_counter())) * 1000.0, 1)


def _describe(file_path: Path, sheet: str | None = None) -> list[dict]:
    """读取首行确定每列名称与推断类型（用 DuckDB 原生 API，不依赖 pandas）。

    优先复用已建 parquet 缓存（读列元数据，更快且类型确定）；无缓存时用
    read_xlsx LIMIT 0 探测（只读表头，几乎不扫数据）。
    """
    con = duckdb.connect()
    tmp_files: list[Path] = []
    try:
        _load_excel(con)
        # 1) 命中缓存（普通类型 或 全文本语义）时直接读 parquet 元数据
        for all_varchar in (False, True):
            if file_path.suffix.lower() != ".csv" and _cached_parquet_exists(
                file_path, sheet, "data", all_varchar
            ):
                parquet = _cached_parquet_path(_source_cache_signature(file_path, sheet, all_varchar))
                cur = con.execute(f"SELECT * FROM read_parquet('{_sql_path(parquet)}') LIMIT 0")
                return [{"name": str(d[0]), "type": str(d[1])} for d in cur.description]
        # 2) 无缓存：read_xlsx LIMIT 0 只读表头
        cur = con.execute(f"SELECT * FROM {_build_source_expr(file_path, sheet, tmp_out=tmp_files)} LIMIT 0")
        return [{"name": str(d[0]), "type": str(d[1])} for d in cur.description]
    finally:
        con.close()
        _cleanup_tmp_files(tmp_files)


def _describe_doris(conn: dict, db: str, table: str) -> list[dict]:
    """返回 Doris 一张表的列元数据（复用共享连接 + 全局锁，避免重复建连）。

    数据源走 information_schema.columns（而非 DESCRIBE / SHOW FULL COLUMNS）：
      - DESCRIBE 只给 column_name/column_type/null，且 key/default/extra 全为 None，
        无注释列；SHOW FULL COLUMNS/SHOW COLUMNS FROM 会被 DuckDB 本地解析器拦截，
        无法下推 Doris。实测 information_schema.columns 返回 24 列，COLUMN_COMMENT
        每列都有中文注释、COLUMN_KEY 标 UNI（主键）或空串、IS_NULLABLE 标 YES/NO、
        COLUMN_TYPE 保留 Doris 真实类型名（如 decimalv3(18,4)）。
    """
    with _SHARED_LOCK:
        con = _get_shared_con()   # 已加载 mysql 扩展 + ATTACH 缓存
        conn_key = (conn["host"], conn["port"], conn["user"], conn["password"])
        attach_alias = _DORIS_ATTACHED.get(conn_key)
        if attach_alias is None:
            attach_alias = f"__doris_{len(_DORIS_ATTACHED) + 1}"
            con.execute(_doris_attach_sql(conn, attach_alias))
            _DORIS_ATTACHED[conn_key] = attach_alias
        # db/table 作字符串字面量进 WHERE（表名可能含点，用 = 精确匹配避免 LIKE 歧义）
        dlit = db.replace("'", "''")
        tlit = table.replace("'", "''")
        cur = con.execute(
            f"SELECT column_name, column_type, column_key, is_nullable, column_comment "
            f"FROM {attach_alias}.information_schema.columns "
            f"WHERE table_schema = '{dlit}' AND table_name = '{tlit}' "
            f"ORDER BY ordinal_position"
        )
        out = []
        for name, ctype, ckey, nullable, comment in cur.fetchall():
            out.append({
                "name": str(name),
                "type": str(ctype),
                "is_key": bool(ckey and str(ckey).strip() != ""),
                "comment": (str(comment).strip() if comment is not None else "") or None,
            })
        return out


def _doris_list(conn: dict) -> dict:
    """连接 Doris 并返回 {dbs: [...], tables: {db: [table,...]}}（供前端选表）。

    实现：ATTACH 不带 db（暴露所有库为 schema），SHOW ALL TABLES 一次拿到
    (attach, schema, table, 列名数组, 类型数组)，据此聚合出库→表清单。
    """
    with _SHARED_LOCK:
        con = _get_shared_con()   # 已加载 mysql 扩展；ATTACH 缓存复用
        conn_key = (conn["host"], conn["port"], conn["user"], conn["password"])
        attach_alias = _DORIS_ATTACHED.get(conn_key)
        if attach_alias is None:
            attach_alias = f"__doris_{len(_DORIS_ATTACHED) + 1}"
            con.execute(_doris_attach_sql(conn, attach_alias))
            _DORIS_ATTACHED[conn_key] = attach_alias
        rows = con.execute("SHOW ALL TABLES").fetchall()
        dbs: set[str] = set()
        tables: dict[str, list[str]] = {}
        for r in rows:
            # (attach, schema, table, [cols], [types], is_column?)
            if len(r) < 3:
                continue
            schema = str(r[1])
            tname = str(r[2])
            dbs.add(schema)
            tables.setdefault(schema, []).append(tname)
        for lst in tables.values():
            lst.sort(key=str.lower)
        return {"dbs": sorted(dbs, key=str.lower), "tables": tables}


class DbProbeRequest(BaseModel):
    conn: DbConn | None = None


@app.post("/api/db_probe")
async def db_probe(payload: DbProbeRequest) -> dict:
    """测试数据库连接并列出库/表（供前端「连接 Doris」表单校验与选表）。

    只读操作（READ_ONLY）。失败时返回 HTTP 400 + 可读错误信息（认证/网络原因）。
    """
    conn = payload.conn
    if not conn or not (conn.host or "").strip():
        raise HTTPException(status_code=400, detail="请填写数据库主机")
    conn_dict = {
        "host": (conn.host or "").strip() or "127.0.0.1",
        "port": int(conn.port or 9030),
        "user": (conn.user or "").strip() or "root",
        "password": conn.password or "",
        "db": (conn.db or "").strip(),
    }
    try:
        listing = await asyncio.to_thread(_doris_list, conn_dict)
        return {
            "ok": True,
            "dbs": listing["dbs"],
            "tables": listing["tables"],
            "server": f"{conn_dict['host']}:{conn_dict['port']}",
        }
    except Exception as e:
        msg = str(e)
        if "access denied" in msg.lower() or "authentication" in msg.lower():
            raise HTTPException(status_code=400, detail=f"认证失败，请检查用户名/密码: {msg[:200]}")
        if "connect" in msg.lower() or "timeout" in msg.lower():
            raise HTTPException(status_code=400, detail=f"无法连接数据库: {msg[:200]}")
        raise HTTPException(status_code=400, detail=f"连接失败: {msg[:200]}")


def _jsonable_value(v):
    """把单个 DuckDB 结果值转成可 JSON 序列化的 Python 值。"""
    import datetime as _dt
    import decimal as _dec

    if v is None:
        return None
    # 注意：datetime 是 date 的子类，必须先判断 datetime
    if isinstance(v, _dt.datetime):
        # 把 ISO 的 T 分隔符换成空格，去掉微秒，如 2026-08-03 15:45:27
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, _dt.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, _dt.time):
        return v.strftime("%H:%M:%S")
    if isinstance(v, _dec.Decimal):
        return float(v)
    if isinstance(v, (bool, int, float, str)):
        return v
    return str(v)


def _jsonable_rows(columns: list[str], rows: list[tuple]) -> list[dict]:
    """把 (列名, 行元组列表) 转成前端需要的 dict 列表。"""
    out: list[dict] = []
    for row in rows:
        out.append({col: _jsonable_value(v) for col, v in zip(columns, row)})
    return out


# ---------- 请求 / 响应模型 ----------

class CheckPathsRequest(BaseModel):
    paths: list[str] = []


class BrowseRequest(BaseModel):
    path: str | None = None


class OpenRequest(BaseModel):
    path: str
    sheet: str | None = None


class Source(BaseModel):
    """一个数据源 = 本地文件 或 远程数据库表 + SQL 别名。

    kind: "file"（默认，兼容旧数据）| "doris"
    doris 源使用 conn（完整连接配置，前端内联携带）+ db + table。
    """

    alias: str
    path: str = ""          # file 源必填；doris 源为空
    sheet: str | None = None
    kind: str = "file"
    conn: DbConn | None = None   # doris：完整连接配置
    db: str = ""            # doris：数据库名
    table: str = ""         # doris：表名


class DbConn(BaseModel):
    """一个数据库连接配置（前端存 localStorage，查询时随源内联传给后端）。"""

    id: str = ""
    name: str = ""
    host: str = ""
    port: int = 9030
    user: str = ""
    password: str = ""
    db: str = ""


class QueryRequest(BaseModel):
    sources: list[Source]
    sql: str = ""       # /api/query 必填（内部校验）；/api/describe 复用本模型但不读 sql，给默认空串防 422
    limit: int = DEFAULT_LIMIT


class ExportRequest(BaseModel):
    columns: list[str]
    rows: list[dict]
    filename: str = "query_result"


# ---------- 接口 ----------

@app.get("/api/health")
async def health() -> dict:
    return {"status": "ok"}


class StoreGetRequest(BaseModel):
    keys: list[str]


class StoreSetRequest(BaseModel):
    kvs: dict[str, str | None]   # key -> JSON 字符串；值为 None 表示删除该 key


@app.post("/api/store/get")
async def store_get(payload: StoreGetRequest) -> dict:
    """批量读取持久化 key（返回 {key: JSON字符串}；缺省 key 不含于结果）。"""
    out: dict[str, str] = {}
    for k in payload.keys:
        v = _store_get(k)
        if v is not None:
            out[k] = v
    return {"values": out}


@app.post("/api/store/set")
async def store_set(payload: StoreSetRequest) -> dict:
    """批量写入/删除持久化 key；value 为 None 视为删除。"""
    for k, v in payload.kvs.items():
        if v is None:
            _store_remove(k)
        else:
            _store_set(k, v)
    return {"ok": True}


@app.post("/api/pick_file")
async def pick_file() -> dict:
    """弹起 Windows 原生「选择文件」对话框，返回所选文件的真实路径。

    后端进程运行在本机，因此能调用系统对话框；浏览器本身拿不到拖拽文件的绝对路径。
    tkinter 对话框会阻塞到用户选择/取消，故放到线程池执行，避免卡住事件循环。
    """
    path = await asyncio.to_thread(_pick_file_via_tkinter)
    if not path:
        return {"path": None, "cancelled": True}
    return {"path": path, "cancelled": False}


@app.post("/api/browse")
async def browse(payload: BrowseRequest) -> dict:
    """浏览本地文件系统：列出盘符、子目录与 Excel 文件（不复制文件）。"""
    return _list_directory(payload.path or "")


@app.post("/api/check_paths")
async def check_paths(payload: CheckPathsRequest) -> dict:
    """批量检查路径是否仍是存在的文件（供「最近打开」标记被移动/删除的文件）。

    只做 stat 级校验（不读内容），用于前端把失效历史条目标红。返回
    {exists: {path: bool}}；路径无需在 ALLOWED_EXTENSIONS 内（历史里可能
    出现过但文件已改名）。
    """
    result: dict[str, bool] = {}
    for raw in payload.paths[:200]:     # 限制单次检查数量，防御异常请求
        p = Path(str(raw).strip().strip('"').strip("'")).expanduser()
        try:
            result[raw] = p.is_file()
        except OSError:
            result[raw] = False
    return {"exists": result}


@app.post("/api/open")
async def open_excel(payload: OpenRequest) -> dict:
    """按路径引用打开 Excel，返回字段列表与工作表列表（不复制文件）。"""
    t_start = time.perf_counter()
    file_path = _resolve_path(payload.path)
    t_resolve = time.perf_counter()
    # 先启动后台预构建（读全表建缓存，是大文件最耗时的部分），与前台 schema 探测/
    # sheet 列表读取并行，缩短「选文件 → 字段可用」的等待。
    threading.Thread(
        target=_prebuild_cache,
        args=(file_path, payload.sheet, "data"),
        daemon=True,
    ).start()
    # schema 探测（LIMIT 0 只读首行）与 sheet 列表（zip 读 workbook.xml）都很快，
    # 放进线程池避免阻塞事件循环。
    columns = await asyncio.to_thread(_describe, file_path, payload.sheet)
    t_describe = time.perf_counter()
    sheets = await asyncio.to_thread(_list_sheets, file_path)
    t_sheets = time.perf_counter()
    return {
        "path": str(file_path),
        "filename": file_path.name,
        "columns": columns,
        "sheets": sheets,
        # 各阶段耗时，供前端在选择文件后连续展示每一步的用时。
        "timings": {
            "resolve_ms": round((t_resolve - t_start) * 1000.0, 1),
            "describe_ms": round((t_describe - t_resolve) * 1000.0, 1),
            "sheets_ms": round((t_sheets - t_describe) * 1000.0, 1),
            "total_ms": round((t_sheets - t_start) * 1000.0, 1),
        },
    }


def _prepare_sources(payload_sources: list[Source]) -> list[dict]:
    """校验并解析多数据源：返回 [{alias, kind, ...}]."""
    if not payload_sources:
        raise HTTPException(status_code=400, detail="至少需要一个数据源")

    seen_aliases: set[str] = set()
    prepared: list[dict] = []
    for src in payload_sources:
        alias = _validate_alias(src.alias)
        if alias.lower() in seen_aliases:
            raise HTTPException(status_code=400, detail=f"别名重复: {alias}")
        seen_aliases.add(alias.lower())

        kind = (src.kind or "file").strip().lower()
        if kind == "doris":
            # 数据库源（连接级）：只校验连接配置；db/table 可选（连接可访问全部库表）。
            # 用户可在 SQL 里显式写 `FROM 别名.库.表`，或由字段区当前选中的库表作为
            # `FROM 别名` 的默认表；两者都不给时后端在执行前报错提示。
            conn = src.conn
            if not conn or not (conn.host or "").strip():
                raise HTTPException(status_code=400, detail=f"{alias}: 未指定主机")
            dbname = (src.db or conn.db or "").strip()
            prepared.append({
                "alias": alias,
                "kind": "doris",
                "conn": {
                    "host": (conn.host or "").strip() or "127.0.0.1",
                    "port": int(conn.port or 9030),
                    "user": (conn.user or "").strip() or "root",
                    "password": conn.password or "",
                    "db": (conn.db or "").strip(),
                },
                "db": dbname,
                "table": str(src.table or "").strip(),
            })
        else:
            # 文件源：现有逻辑
            _sheet_literal(src.sheet)
            resolved = _resolve_path(src.path)
            prepared.append({
                "alias": alias,
                "kind": "file",
                "path": src.path,
                "resolved": resolved,
                "sheet": (src.sheet or "").strip() or None,
            })
    return prepared


def _code_tags(sql: str) -> list[str]:
    """返回与 SQL 等长的类型标记：'c'（代码）/ 's'（字符串/引号标识符）/ 'x'（注释）。

    供 _finalize_doris_sql 做「尾部裁剪到最后一个真实代码字符」用，避免把 SQL
    末尾的字符串字面量（如 'a--b'）误当注释裁掉。
    """
    tags = ["c"] * len(sql)
    i, n = 0, len(sql)
    while i < n:
        c = sql[i]
        if c in ("'", '"', "`"):
            q = c
            j = i + 1
            while j < n:
                if sql[j] == "\\":
                    j += 2
                    continue
                if sql[j] == q:
                    break
                j += 1
            for k in range(i, min(j + 1, n)):
                tags[k] = "s"
            i = j + 1
        elif c == "-" and i + 1 < n and sql[i + 1] == "-":
            j = sql.find("\n", i)
            j = n if j == -1 else j
            for k in range(i, j):
                tags[k] = "x"
            i = j
        elif c == "#":
            j = sql.find("\n", i)
            j = n if j == -1 else j
            for k in range(i, j):
                tags[k] = "x"
            i = j
        elif c == "/" and i + 1 < n and sql[i + 1] == "*":
            j = sql.find("*/", i + 2)
            j = n - 2 if j == -1 else j + 2
            for k in range(i, j):
                tags[k] = "x"
            i = j
        else:
            i += 1
    return tags


def _finalize_doris_sql(sql: str, limit: int) -> str:
    """Doris 源执行的最终 SQL：剥尾部注释/分号；若顶层无 LIMIT 则追加（让 LIMIT 下推）。

    视图包装会把 LIMIT 吃掉导致全表拉取（见性能诊断），Doris 源必须直接执行。
    基于 _code_tags 定位最后一个真实「代码」字符：尾部注释/分号剥掉，
    字符串字面量（'a--b'）里的内容保留；子查询里的 LIMIT 不影响顶层判断。
    """
    tags = _code_tags(sql)
    # 最后一个「内容」字符（代码 c 或字符串 s）的位置；其后只允许空白/分号/注释。
    # 注意字符串 s 也是合法内容：SQL 以字面量结尾（如 WHERE name = 'a--b'）时必须保留。
    last_content = -1
    for i, t in enumerate(tags):
        if t in ("c", "s"):
            last_content = i
    if last_content < 0:
        return sql
    core = sql[: last_content + 1].rstrip()
    # 剥末尾分号（允许 ; ; 连续 / 分号+空白）
    while core.endswith(";"):
        core = core[:-1].rstrip()
    # 顶层末尾是否已有 LIMIT（跨空白；允许 LIMIT n / LIMIT n OFFSET m / LIMIT ALL）
    if re.search(r"\blimit\b\s+(\d+|all)(\s+offset\s+\d+)?\s*$", core, re.IGNORECASE):
        return core
    return f"{core}\nLIMIT {limit}"


def _rewrite_doris_sql(sql: str, sources: list[dict]) -> str:
    """把用户 SQL 里的 doris 别名（表引用位置）展开成三段路径 + AS 别名。

    mysql_scanner 扩展在「视图包装后聚合」有绑定 bug（见 P0 实测），但直接用
    mysql 表（三段引用）完全正常。因此 doris 源不建视图，而是在执行前把
    SQL 中 FROM / JOIN /, 后的表引用位置展开：

      显式三段（连接级源，推荐）：
          FROM d1.库.表          → FROM __doris_1."库"."表" AS d1
          JOIN d1.库.表 ON ...   → JOIN __doris_1."库"."表" AS d1 ON ...
      裸别名（有默认库表时，向后兼容）：
          FROM d1                → FROM __doris_1."默认库"."默认表" AS d1
      裸别名但无默认库表：报错，提示改用显式三段或先在字段区选库表。

    列引用（d1.PERIOD）保留不动（d1 仍是表别名）。只处理 doris 源；file 源
    别名仍是视图，不展开。
    """
    def _unquote_ident(tok: str) -> str:
        tok = tok.strip()
        if len(tok) >= 2 and tok[0] == '"' and tok[-1] == '"':
            return tok[1:-1].replace('""', '"')
        return tok

    # alias -> {attach, db, table}；db/table 可能为空（连接级源未选默认表）
    doris_info: dict[str, dict] = {}
    for s in sources:
        if s.get("kind") == "doris" and s.get("_attach_alias"):
            doris_info[s["alias"]] = {
                "attach": s["_attach_alias"],
                "db": (s.get("db") or "").strip(),
                "table": (s.get("table") or "").strip(),
            }
    if not doris_info:
        return sql

    # 标识符：裸 [字母数字下划线]+ 或 "..."双引号包裹
    ident = r'(?:"(?:[^"]|"")*"|[A-Za-z0-9_]+)'
    # 表引用后常见的子句关键字：裸表别名捕获时不能把这些字吞进去
    reserved_trail = (
        "where|on|group|order|limit|offset|having|union|except|intersect|"
        "left|right|inner|outer|full|cross|join|natural|as|using|qualify|"
        "window|fetch|first|next|rows|select|from|with|asc|desc|all|distinct|"
        "option|pivot|unpivot|returning|into|then|else|end|over|partition|range"
    )

    def _tbl_alias(m, as_group, bare_group, default):
        """取用户跟在表引用后的手动别名（AS x / 裸 x）；无则用默认连接别名。"""
        tok = None
        if as_group and m.group(as_group):
            tok = m.group(as_group)
        elif bare_group and m.group(bare_group):
            tok = m.group(bare_group)
        return _unquote_ident(tok) if tok else default

    # 按别名长短降序处理，避免短别名先替换吃掉长别名前缀
    for alias in sorted(doris_info, key=len, reverse=True):
        info = doris_info[alias]
        a = re.escape(alias)

        # 1) 显式三段：FROM/JOIN/, 后 别名 . 库 . 表，可选手动表别名（AS x / 裸 x）
        #    分组：1=kw 2=库 3=表 4=AS别名 5=裸别名
        explicit = re.compile(
            rf"(?<!\w)(FROM|JOIN|,)\s+{a}\s*\.\s*({ident})\s*\.\s*({ident})"
            rf"(?:\s+AS\s+({ident})|\s+(?!(?:{reserved_trail})\b)({ident}))?(?![\w.])",
            re.IGNORECASE,
        )
        def _ex_sub(m, _info=info, _alias=alias):
            kw = m.group(1)
            db = _unquote_ident(m.group(2))
            tbl = _unquote_ident(m.group(3))
            talias = _tbl_alias(m, 4, 5, _alias)
            dbq = '"' + db.replace('"', '""') + '"'
            tblq = '"' + tbl.replace('"', '""') + '"'
            return f"{kw} {_info['attach']}.{dbq}.{tblq} AS {talias}"
        sql = explicit.sub(_ex_sub, sql)

        # 2) 裸别名：FROM/JOIN/, 后 别名，可选手动表别名（AS x / 裸 x）
        #    分组：1=kw 2=AS别名 3=裸别名
        bare = re.compile(
            rf"(?<!\w)(FROM|JOIN|,)\s+{a}"
            rf"(?:\s+AS\s+({ident})|\s+(?!(?:{reserved_trail})\b)({ident}))?(?![\w.])",
            re.IGNORECASE,
        )
        if info["db"] and info["table"]:
            dbq = '"' + info["db"].replace('"', '""') + '"'
            tblq = '"' + info["table"].replace('"', '""') + '"'
            path = f"{info['attach']}.{dbq}.{tblq}"
            def _bare_sub(m, _path=path, _alias=alias):
                return f"{m.group(1)} {_path} AS {_tbl_alias(m, 2, 3, _alias)}"
            sql = bare.sub(_bare_sub, sql)
        elif bare.search(sql):
            raise HTTPException(
                status_code=400,
                detail=(
                    f"别名 {alias} 未指定库表：请在 SQL 写 FROM {alias}.库.表，"
                    f"或在数据字段区为 {alias} 选择库表"
                ),
            )
    return sql


def _doris_attach_sql(conn: dict, attach_alias: str, db: str | None = None) -> str:
    """构造 Doris ATTACH 语句（TYPE mysql, READ_ONLY）。

    db 缺省时用连接配置里的默认库。注意：为支持同时引用多库的表（一次 ATTACH 暴露
    该连接的所有库为 schema），实际实现里始终 ATTACH 不带 db；db 参数保留用于未来
    指定默认库的场景。
    """
    host = conn.get("host") or "127.0.0.1"
    port = conn.get("port") or 9030
    user = conn.get("user") or "root"
    pw = conn.get("password") or ""
    parts = [f"host={host}", f"port={port}", f"user={user}"]
    if pw:
        parts.append(f"passwd={pw}")
    # 注意：不带 db，让 Doris 所有库都映射为 DuckDB schema
    conn_str = " ".join(parts)
    # 别名进 SQL 必须加双引号防注入（ATTACH 的别名是标识符）
    quoted = '"' + attach_alias.replace('"', '""') + '"'
    return f"ATTACH '{conn_str}' AS {quoted} (TYPE mysql, READ_ONLY)"


def _register_sources(con: duckdb.DuckDBPyConnection, sources: list[dict], all_varchar: bool = False,
                      tmp_files: list[Path] | None = None) -> tuple[bool, dict[str, str]]:
    """把每个数据源注册成名字 = 别名的视图。

    返回 (是否新建了缓存, {别名: 缓存状态})。cache_status 取值：
      - "hit"     命中已有 parquet 缓存
      - "built"   本次扫描源表并新建缓存
      - "csv"     该源是 CSV，未用缓存
      - "direct"  未启用缓存（直接读源，如缓存构建失败后的回退）

    CSV 转码产生的临时文件会追加到 tmp_files（由调用方在 fetchall 后统一清理）。
    """
    built_any = False
    status: dict[str, str] = {}

    for s in sources:
        quoted = '"' + s["alias"].replace('"', '""') + '"'
        expr = None
        alias_status = "direct"

        # 数据库源（Doris）：不建视图（mysql 扩展视图+聚合有绑定 bug，见 P0 实测），
        # 只负责 ATTACH；同连接只 ATTACH 一次（进程级缓存），SQL 里别名由
        # _rewrite_doris_sql 展开。ATTACH 在共享连接上进行（调用方已持 _SHARED_LOCK）。
        if s.get("kind") == "doris":
            conn = s["conn"]
            conn_key = (conn["host"], conn["port"], conn["user"], conn["password"])
            attach_alias = _DORIS_ATTACHED.get(conn_key)
            if attach_alias is None:
                attach_alias = f"__doris_{len(_DORIS_ATTACHED) + 1}"
                con.execute(_doris_attach_sql(conn, attach_alias))
                _DORIS_ATTACHED[conn_key] = attach_alias
            s["_attach_alias"] = attach_alias
            alias_status = "doris"
            status[s["alias"]] = alias_status
            continue
        elif s["resolved"].suffix.lower() == ".csv":
            # CSV 不参与 parquet 缓存（本就快），直接走 read_csv 表达式，
            # 避免落入「parquet=None → read_parquet('None')」的畸形兜底分支。
            expr = _build_source_expr(s["resolved"], s["sheet"], all_varchar=False, tmp_out=tmp_files)
            alias_status = "csv"
        elif not all_varchar:
            try:
                parquet, built = _parquet_cache_load(con, s["resolved"], s["sheet"], s["alias"], all_varchar)
                if parquet is not None:
                    expr = f"read_parquet('{_sql_path(parquet)}')"
                    alias_status = "built" if built else "hit"
                    built_any = built_any or built
            except Exception:
                # 缓存构建失败（通常是首次扫描命中单元格转换错误）→ 走正常表达式，
                # 让上层 all_varchar 回退机制接管
                expr = None
                alias_status = "direct"

        if expr is None:
            if all_varchar:
                # 兜底重读：走全文本语义。若已有该语义缓存则命中；否则新建。
                try:
                    parquet, built = _parquet_cache_load(con, s["resolved"], s["sheet"], s["alias"], all_varchar=True)
                    expr = f"read_parquet('{_sql_path(parquet)}')"
                    alias_status = "built" if built else "hit"
                    built_any = built_any or built
                except Exception:
                    expr = _build_source_expr(s["resolved"], s["sheet"], all_varchar=True, tmp_out=tmp_files)
                    alias_status = "direct"
            else:
                expr = _build_source_expr(s["resolved"], s["sheet"], all_varchar=False, tmp_out=tmp_files)
                alias_status = "direct"

        status[s["alias"]] = alias_status
        con.execute(f"CREATE OR REPLACE VIEW {quoted} AS SELECT * FROM {expr}")

    return built_any, status


@app.post("/api/describe")
async def describe_sources(payload: QueryRequest) -> dict:
    """批量返回每个数据源的字段与工作表列表（供多源模式一次加载）。"""
    sources = _prepare_sources(payload.sources)
    result = []
    for s in sources:
        if s.get("kind") == "doris":
            # 数据库源：列来自 ATTACH 后 DESCRIBE，无 sheet/无预构建
            columns = await asyncio.to_thread(
                _describe_doris, s["conn"], s["db"], s["table"]
            )
            result.append({
                "alias": s["alias"],
                "kind": "doris",
                "path": "",
                "filename": f"{s['db']}.{s['table']}",
                "sheet": None,
                "columns": columns,
                "sheets": [],
            })
            continue
        columns = _describe(s["resolved"], s["sheet"])
        sheets = _list_sheets(s["resolved"])
        # 选完源后，后台预构建缓存（多源模式每个源都要）
        threading.Thread(
            target=_prebuild_cache,
            args=(s["resolved"], s["sheet"], s["alias"]),
            daemon=True,
        ).start()
        result.append({
            "alias": s["alias"],
            "kind": "file",
            "path": str(s["resolved"]),
            "filename": s["resolved"].name,
            "sheet": s["sheet"],
            "columns": columns,
            "sheets": sheets,
        })
    return {"sources": result}


class PrebuildRequest(BaseModel):
    path: str
    sheet: str | None = None


@app.post("/api/prebuild_status")
async def prebuild_status(payload: PrebuildRequest) -> dict:
    """查询某文件的后台预构建状态：running / done / error / none（未触发）。

    用于前端在选完文件后展示「正在后台预处理…」，直到就绪。
    """
    try:
        resolved = _resolve_path(payload.path)
    except HTTPException:
        return {"status": "none"}
    key = _prebuild_key(resolved, payload.sheet)
    with _PREBUILD_LOCK:
        st = _PREBUILD_STATUS.get(key)
    if not st:
        return {"status": "none"}
    # 连续展示：status 之外带出已耗时（进行中按当前时刻，结束时按最终耗时）。
    started = st.get("started")
    if st["status"] == "running" and started is not None:
        elapsed_ms = round((time.perf_counter() - started) * 1000.0, 1)
    else:
        elapsed_ms = st.get("elapsed_ms")
    return {"status": st["status"], "elapsed_ms": elapsed_ms}


# ---------- 异步查询任务（真实进度百分比） ----------
# DuckDB 的查询是同步阻塞的，进度只能靠另一线程轮询 con.query_progress()。
# 因此把 /api/query 改为「提交任务」：立即返回 task_id，后台线程跑查询、
# 轮询线程采进度，前端通过 /api/progress 拿真实百分比与最终结果。

_TASKS: dict[str, dict] = {}
_TASKS_LOCK = threading.Lock()


def _run_query_task(task_id: str, sources: list[dict], sql: str, limit: int) -> None:
    """后台线程：执行查询，期间由进度线程轮询 query_progress 写回进度。

    分两个阶段：
      - prepare：加载 excel 扩展 + 视图注册（含 GBK 转码）+ DuckDB 嗅探，
        此阶段无真实进度，前端用不确定动画提示「正在加载数据」；
      - run：进入真实查询（execute/fetchall），query_progress 给出 0~100 百分比。
    """
    task = _TASKS[task_id]
    t_start = time.perf_counter()   # 全流程计时起点：含读 Excel / 缓存构建 + SQL 执行
    has_doris = any(s.get("kind") == "doris" for s in sources)
    con = duckdb.connect()
    tmp_files: list[Path] = []
    if has_doris:
        # Doris 源：走进程级共享连接（复用 ATTACH，省掉 ~1.8s 一次性开销）。
        # 整个查询持 _SHARED_LOCK 串行，避免共享连接上的视图/状态被并发踩踏。
        _SHARED_LOCK.acquire()
    try:
        # 关闭进度条打印，并让 progress_bar 从 0ms 起算进度（否则前 2s 内
        # query_progress() 返回 -1.0，拿不到真实百分比）
        if has_doris:
            con = _get_shared_con()
        else:
            con.execute("SET enable_progress_bar = true")
            con.execute("SET enable_progress_bar_print = false")
            con.execute("SET progress_bar_time = 0")
            _load_excel(con)

        # 语句类型决定执行方式：
        #   - SELECT / WITH：注册成临时视图再套 LIMIT（对行注释/块注释/末尾分号都健壮，
        #     避免把 SQL 直接塞进 "SELECT * FROM (...)" 导致行尾注释注释掉右括号、分号报错）
        #   - DESCRIBE / SHOW / SUMMARIZE / PRAGMA：直接原样执行（同样容忍注释/分号）
        stmt_kind = _leading_sql_keyword(sql)
        tmp_view = "__dq_user_query"
        used = {s["alias"].lower() for s in sources}
        n = 1
        while tmp_view.lower() in used:  # 避免与用户数据源别名冲突
            n += 1
            tmp_view = f"__dq_user_query_{n}"
        qtmp = '"' + tmp_view.replace('"', '""') + '"'

        # 混合类型（数字列夹杂文本）文件在普通类型扫描下必然报「单元格转换错误」，
        # 首次查询会失败再兜底成全文本并建缓存。若所有 xlsx 源都已存在「全文本语义」
        # 缓存（即此前已判定为混合类型），直接走全文本缓存，跳过注定失败的普通扫描，
        # 避免每次白读一遍整表。数据库源（doris）不参与此判断（无文件）。
        _all_have_varchar_cache = bool(sources) and all(
            s.get("kind") == "doris"
            or s["resolved"].suffix.lower() == ".csv"
            or _cached_parquet_exists(s["resolved"], s["sheet"], s["alias"], all_varchar=True)
            for s in sources
        )
        phases = (True,) if _all_have_varchar_cache else (False, True)

        for all_varchar in phases:
            # 第一次按正常类型推断扫描；若命中「数字列里夹杂文本」的单元格转换错误，
            # 把 xlsx 数据源整表按文本重读再执行一次（保留所有值，不丢数据）。
            built_any, cache_status = _register_sources(con, sources, all_varchar, tmp_files)

            # 准备阶段结束，进入真实查询阶段
            with _TASKS_LOCK:
                task["phase"] = "run"
                task["progress"] = 0.0

            t0 = time.perf_counter()   # 仅 SQL 执行段的计时起点

            # 进度轮询线程：查询阻塞期间持续采样 query_progress()
            stop_poll = threading.Event()

            def _poll() -> None:
                while not stop_poll.is_set():
                    try:
                        p = con.query_progress()
                        if isinstance(p, (int, float)):
                            if 0.0 <= p <= 100.0:
                                task["progress"] = max(0.0, min(100.0, float(p)))
                    except Exception:
                        pass
                    time.sleep(0.05)

            poller = threading.Thread(target=_poll, daemon=True)
            poller.start()

            try:
                # doris 源：把用户 SQL 里的别名（表引用位置）展开成三段路径 + AS 别名，
                # 因为 mysql 扩展视图+聚合有绑定 bug，doris 源不建视图（见 _rewrite_doris_sql）。
                # 只允许 SELECT/WITH 时重写；元数据语句（describe 等）按原样跑。
                has_doris = any(s.get("kind") == "doris" for s in sources)
                exec_sql = _rewrite_doris_sql(sql, sources) if stmt_kind not in ("describe", "show", "summarize", "pragma") else sql
                if stmt_kind in ("describe", "show", "summarize", "pragma"):
                    # 元数据语句：直接执行（不容忍被装进子查询/视图）
                    cur = con.execute(exec_sql)
                elif has_doris:
                    # Doris 源：直接执行 + 追加 LIMIT（视图包装会把 LIMIT 吃掉导致
                    # 全表拉取，见 P0 性能诊断；直接执行让 LIMIT/谓词下推到 Doris）。
                    cur = con.execute(_finalize_doris_sql(exec_sql, limit))
                else:
                    con.execute(f"CREATE OR REPLACE VIEW {qtmp} AS {exec_sql}")
                    cur = con.execute(f"SELECT * FROM {qtmp} LIMIT {limit}")
                columns = [str(d[0]) for d in cur.description]
                raw_rows = cur.fetchall()
                elapsed_ms = (time.perf_counter() - t0) * 1000.0       # SQL 执行段
                total_ms = (time.perf_counter() - t_start) * 1000.0    # 含准备段的全流程
                prepare_ms = total_ms - elapsed_ms                     # 读文件/缓存构建段
                break  # 查询成功，跳出回退循环
            except Exception as e:
                # 首次扫描触发「数字列里夹杂文本」的转换错误时，回退为全文本重读
                if not all_varchar and _is_cell_conversion_error(e):
                    continue
                raise
            finally:
                stop_poll.set()
                poller.join(timeout=1)
        else:
            # 逻辑上不可达（异常会向上抛出）；保留兜底，便于静态分析/未来改动
            raise RuntimeError("查询未执行")

        rows = _jsonable_rows(columns, raw_rows)
        # 汇总缓存状态给前端展示：任一源命中/新建缓存即标记
        cache_state = "none"
        if built_any:
            cache_state = "built"
        elif any(v == "hit" for v in cache_status.values()):
            cache_state = "hit"
        with _TASKS_LOCK:
            task["status"] = "done"
            task["phase"] = "done"
            task["progress"] = 100.0
            task["result"] = {
                "columns": columns,
                "rows": rows,
                "row_count": len(rows),
                "limit": limit,
                "elapsed_ms": round(elapsed_ms, 2),
                "prepare_ms": round(prepare_ms, 2),
                "total_ms": round(total_ms, 2),
                "cache": cache_state,
                "cache_status": cache_status,
            }
    except HTTPException as e:
        with _TASKS_LOCK:
            task["status"] = "error"
            task["phase"] = "error"
            task["error"] = str(getattr(e, "detail", "查询失败"))
    except Exception as e:  # noqa: BLE001
        with _TASKS_LOCK:
            task["status"] = "error"
            task["phase"] = "error"
            task["error"] = f"SQL 执行出错: {e}"
    finally:
        if has_doris:
            # 共享连接由应用退出时统一关闭（_close_shared_con），这里只释放锁
            try:
                _SHARED_LOCK.release()
            except Exception:
                pass
        else:
            try:
                con.close()
            except Exception:
                pass
        _cleanup_tmp_files(tmp_files)


@app.post("/api/query")
async def query_excel(payload: QueryRequest) -> dict:
    """提交查询任务，立即返回 {task_id}；进度与结果走 /api/progress。"""
    sources = _prepare_sources(payload.sources)

    sql = (payload.sql or "").strip()
    if not sql:
        raise HTTPException(status_code=400, detail="SQL 不能为空")

    first_word = _leading_sql_keyword(sql)
    allowed = {"select", "with", "describe", "show", "summarize", "pragma"}
    if first_word not in allowed:
        raise HTTPException(status_code=400, detail="仅允许 SELECT / WITH / DESCRIBE / SHOW 等只读查询")

    _assert_no_forbidden_tables(sql)

    limit = payload.limit
    if not isinstance(limit, int) or limit <= 0:
        limit = DEFAULT_LIMIT
    limit = min(limit, MAX_RESULT_ROWS)

    task_id = uuid.uuid4().hex
    with _TASKS_LOCK:
        _TASKS[task_id] = {
            "status": "running",
            "phase": "prepare",
            "progress": 0.0,
            "result": None,
            "error": None,
        }

    t = threading.Thread(
        target=_run_query_task,
        args=(task_id, sources, sql, limit),
        daemon=True,
    )
    t.start()

    return {"task_id": task_id}


class ProgressRequest(BaseModel):
    task_id: str


@app.post("/api/progress")
async def query_progress(payload: ProgressRequest) -> dict:
    """查询任务进度：{ status, progress, result?, error? }。done 时附带 result。"""
    with _TASKS_LOCK:
        task = _TASKS.get(payload.task_id)
    if task is None:
        raise HTTPException(status_code=404, detail="查询任务不存在或已过期")
    status = task["status"]
    out: dict = {
        "status": status,
        "phase": task.get("phase", "run"),
        "progress": round(float(task.get("progress", 0.0)), 2),
    }
    if status == "done":
        out["result"] = task.get("result")
        # 返回后即可清理，避免任务堆积
        with _TASKS_LOCK:
            _TASKS.pop(payload.task_id, None)
    elif status == "error":
        out["error"] = task.get("error")
        with _TASKS_LOCK:
            _TASKS.pop(payload.task_id, None)
    return out


def _export_to_file(columns: list[str], rows: list[dict], save_path: str, fmt: str) -> int:
    """把查询结果写为 CSV 或 XLSX 文件，返回写入行数。"""
    path = Path(save_path)
    if fmt == "xlsx":
        import openpyxl

        wb = openpyxl.Workbook(write_only=True)
        ws = wb.create_sheet(title="查询结果")
        ws.append(columns)
        for r in rows:
            ws.append([r.get(c) for c in columns])
        wb.save(str(path))
        return len(rows)
    # CSV（默认），用 csv 模块自动处理引号转义；BOM 便于 Excel 识别中文
    import csv

    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(columns)
        for r in rows:
            w.writerow([r.get(c) for c in columns])
    return len(rows)


@app.post("/api/export")
async def export_result(payload: ExportRequest) -> dict:
    """弹出「另存为」对话框，把查询结果导出为 Excel(.xlsx) 或 CSV。"""
    columns = payload.columns or []
    rows = payload.rows or []
    if not columns:
        raise HTTPException(status_code=400, detail="没有可导出的列")

    # 让用户选择保存格式与位置（阻塞），放到线程池
    def _choose_and_write() -> dict:
        filetypes = [
            ("Excel 工作簿 (*.xlsx)", "*.xlsx"),
            ("CSV 文件 (*.csv)", "*.csv"),
        ]
        save_path = _pick_save_path(payload.filename or "query_result", filetypes)
        if not save_path:
            return {"saved": False, "cancelled": True, "path": None}

        fmt = ".xlsx" if save_path.lower().endswith(".xlsx") else "csv"
        if save_path.lower().endswith(".csv") is False and not save_path.lower().endswith(".xlsx"):
            # 用户未填扩展名时按文件类型默认
            save_path += ".xlsx" if fmt == "xlsx" else ".csv"

        count = _export_to_file(columns, rows, save_path, fmt)
        return {"saved": True, "cancelled": False, "path": str(save_path), "format": fmt, "rows": count}

    return await asyncio.to_thread(_choose_and_write)


# 挂载前端静态页面
if FRONTEND_DIR.exists():
    app.mount("/", StaticFiles(directory=str(FRONTEND_DIR), html=True), name="frontend")


def find_free_port() -> int:
    """返回一个空闲的本地端口（供桌面启动器使用）。"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


_server_handle: "uvicorn.Server | None" = None


def run_server(host: str = "127.0.0.1", port: int = 8000) -> None:
    """以线程化方式运行 uvicorn（供桌面启动器调用）。

    保存 Server 句柄，供 shutdown() 优雅停止。
    windowed（console=False）打包后 sys.stdout/stderr 为 None，会令 uvicorn
    的日志初始化崩溃，故先补齐到 os.devnull，并关闭其自带日志配置。
    """
    import logging
    import uvicorn

    # windowed exe 里 sys.stdout/stderr 可能为 None → 补齐避免 dictConfig/emit 崩溃
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")

    global _server_handle
    config = uvicorn.Config(
        app,
        host=host,
        port=port,
        log_level="warning",
        log_config=None,   # 跳过 uvicorn 的 dictConfig（其 StreamHandler 依赖 stderr）
    )
    _server_handle = uvicorn.Server(config)
    _server_handle.run()


def shutdown() -> None:
    """优雅停止由 run_server() 启动的后台服务（桌面窗口关闭时调用）。"""
    global _server_handle
    # 关闭进程级共享 Doris 连接（ATTACH 复用连接，随应用退出一起释放）
    _close_shared_con()
    # 关闭 SQLite 持久化连接
    _close_store()
    if _server_handle is not None:
        _server_handle.should_exit = True


if __name__ == "__main__":
    run_server()