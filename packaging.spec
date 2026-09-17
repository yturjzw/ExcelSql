# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：Excel SQL 查询台（onedir 桌面应用）。

入口 desktop.py：
  - 内嵌 FastAPI + DuckDB（含 excel 扩展，离线）后台服务
  - pywebview（WebView2）原生窗口

生成产物：dist/ExcelSqlConsole/ExcelSqlConsole.exe
"""
from PyInstaller.utils.hooks import collect_all

block_cipher = None

# pywebview（winforms/edgechromium 依赖 pythonnet 的 .NET 程序集与动态库）
webview_datas, webview_binaries, webview_hidden = collect_all("webview")

a = Analysis(
    ["desktop.py"],
    pathex=["."],
    binaries=webview_binaries,
    datas=webview_datas + [
        ("frontend", "frontend"),          # 前端静态资源
        ("duckdb_ext", "duckdb_ext"),      # DuckDB excel 扩展（离线）
    ],
    hiddenimports=webview_hidden + [
        "webview.platforms.winforms",
        "webview.platforms.edgechromium",
        "clr_loader",
        "openpyxl",                 # 结果导出 Excel(.xlsx)
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinter.test"],  # 保留 tkinter 本体（文件对话框需要）
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,   # onedir：主程序与依赖分开
    name="ExcelSqlConsole",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,           # 桌面应用：不显示控制台窗口
    icon="assets/app.ico",   # 与网页版 .brand-mark 一致（近黑圆角方块 + 白色表格/箭头）
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="ExcelSqlConsole",
)