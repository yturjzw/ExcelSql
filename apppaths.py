"""可靠的应用数据目录解析。

冻结（PyInstaller）的 windowed 进程里，os.environ 可能缺失 LOCALAPPDATA 等
用户环境变量，导致持久化/日志被错误地回退到 exe 所在目录（读空库、历史"丢失"）。
这里用 Windows 已知文件夹 API 作为兜底，保证拿到真实的 %LOCALAPPDATA%。
"""
from __future__ import annotations

import ctypes
import os
import sys
from pathlib import Path

APP_DIR_NAME = "ExcelSqlConsole"

# FOLDERID_LocalAppData = {F1B32785-6FBA-4FCF-9D55-7B8E7F157091}
class _GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", ctypes.c_uint32),
        ("Data2", ctypes.c_uint16),
        ("Data3", ctypes.c_uint16),
        ("Data4", ctypes.c_uint8 * 8),
    ]


_FOLDERID_LOCAL_APPDATA = _GUID(
    0xF1B32785,
    0x6FBA,
    0x4FCF,
    (ctypes.c_uint8 * 8)(0x9D, 0x55, 0x7B, 0x8E, 0x7F, 0x15, 0x70, 0x91),
)


def _win_local_appdata() -> str | None:
    """通过 SHGetKnownFolderPath 取 LOCALAPPDATA（不依赖环境变量）。"""
    try:
        shell32 = ctypes.windll.shell32
        shell32.SHGetKnownFolderPath.argtypes = [
            ctypes.c_void_p,  # REFKNOWNFOLDERID (GUID*)
            ctypes.c_uint32,  # DWORD dwFlags
            ctypes.c_void_p,  # HANDLE hToken
            ctypes.POINTER(ctypes.c_wchar_p),  # PWSTR *ppszPath
        ]
        shell32.SHGetKnownFolderPath.restype = ctypes.c_long  # HRESULT

        buf = ctypes.c_wchar_p()
        hr = shell32.SHGetKnownFolderPath(
            ctypes.byref(_FOLDERID_LOCAL_APPDATA), 0, None, ctypes.byref(buf)
        )
        if hr == 0 and buf.value:
            return str(buf.value)
    except Exception:  # noqa: BLE001
        pass
    return None


def local_appdata() -> str | None:
    """返回真实 LOCALAPPDATA 绝对路径；优先环境变量，缺失时走 Windows API。"""
    v = os.environ.get("LOCALAPPDATA")
    if v:
        return v
    if sys.platform == "win32":
        return _win_local_appdata()
    return None


def app_data_dir() -> Path | None:
    """返回应用数据目录 %LOCALAPPDATA%\\ExcelSqlConsole（自动创建）；不可用返回 None。"""
    base = local_appdata()
    if not base:
        return None
    d = Path(base) / APP_DIR_NAME
    try:
        d.mkdir(parents=True, exist_ok=True)
        return d
    except OSError:
        return None


def exe_dir() -> Path:
    """冻结时为 exe 所在目录；源码运行时为本文件所在目录（项目根）。"""
    if getattr(sys, "frozen", False):
        return Path(os.path.dirname(sys.executable))
    return Path(__file__).resolve().parent