# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：Excel SQL 查询台（onedir 桌面应用）。

入口 desktop.py：
  - 内嵌 FastAPI + DuckDB（含 excel 扩展，离线）后台服务
  - pywebview（gui="qt"，QtWebEngine / PySide6）原生窗口

生成产物：dist/ExcelSqlConsole/ExcelSqlConsole.exe
"""
from PyInstaller.utils.hooks import collect_all
import sys  # noqa: F401   PyInstaller 6.x 不自动注入 sys，spec 内需显式导入
import os

block_cipher = None

# pywebview（qt 后端）
webview_datas, webview_binaries, webview_hidden = collect_all("webview")

# QtWebEngine 原生窗口后端：PySide6 6.x 拆分为多个包，另有 qtpy shim。
# 统一 collect_all，交给 PyInstaller 内置 hook（hook-PySide6 / hook-qtpy）收集
# QtWebEngineProcess.exe、resources、translations 等运行时资源。
_qt_datas, _qt_bins, _qt_hidden = [], [], []
for _pkg in (
    "PySide6",
    "PySide6_Essentials",
    "PySide6_Addons",
    "PySide6_WebEngine",
    "shiboken6",
    "qtpy",
):
    try:
        _d, _b, _h = collect_all(_pkg)
        _qt_datas += _d
        _qt_bins += _b
        _qt_hidden += _h
    except Exception:
        pass

a = Analysis(
    ["desktop.py"],
    pathex=["."],
    binaries=webview_binaries + _qt_bins,
    datas=webview_datas + _qt_datas + [
        ("frontend", "frontend"),          # 前端静态资源
        ("duckdb_ext", "duckdb_ext"),      # DuckDB excel 扩展（离线）
    ],
    hiddenimports=webview_hidden + _qt_hidden + [
        "webview.platforms.qt",
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