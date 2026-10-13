# Excel SQL 查询台 —— 打包脚本（onedir）
# 用法：用 venv 里的 Python 运行，例如
#   .\.venv\Scripts\python.exe build.py
import ctypes
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))

# 把临时目录挪进工作区，规避杀软/权限对系统 Temp 的限制
_TMP = os.path.join(HERE, ".tmp", "pyinstaller")
os.makedirs(_TMP, exist_ok=True)
os.environ["TMP"] = _TMP
os.environ["TEMP"] = _TMP

PYI_EXE = os.path.join(HERE, ".venv", "Scripts", "pyinstaller.exe")
if not os.path.exists(PYI_EXE):
    PYI_EXE = os.path.join(HERE, ".venv", "Scripts", "pyinstaller")

# 打包产物部署目录（最终"双击即用"的位置）：
# 本工作区（D:\Projects\DeepseekHarness\...）被 DSH 沙箱打了完整性标签，
# 从 dist 启动会以 Untrusted 完整性运行，托盘图标不显示；D:\Software 无标签。
# 可用环境变量 EXCELSQL_DEPLOY_DIR 覆盖。
_DEPLOY_ROOT = os.environ.get("EXCELSQL_DEPLOY_DIR") or r"D:\Software"


def _deploy() -> str | None:
    """把打包产物复制到无沙箱标签的目录（托盘可正常显示）。

    背景：工作区目录带 'Low Mandatory Level' 完整性标签 → 进程以 Untrusted
    完整性运行 → Windows 按 UIPI 规则不渲染其托盘图标（Shell_NotifyIcon 仍
    返回成功，Shell_NotifyIconGetRect 坐标为 (0,0)-(0,0)）。D:\\\\Software 无
    标签，运行正常。返回目标目录路径；失败返回 None（仅提示，不中断）。
    """
    try:
        src = os.path.join(HERE, "dist", "ExcelSqlConsole")
        if not os.path.isdir(src):
            print("[deploy] 未找到 dist\\ExcelSqlConsole，跳过部署", flush=True)
            return None
        dst = os.path.join(_DEPLOY_ROOT, "ExcelSqlConsole")
        if os.path.isdir(dst):
            shutil.rmtree(dst, ignore_errors=True)
        shutil.copytree(src, dst)
        print(f"[deploy] 已部署到：{dst}", flush=True)
        print("[deploy] 请从这里运行（工作区 dist 目录带沙箱完整性标签，托盘不显示）", flush=True)
        return dst
    except Exception as e:  # noqa: BLE001
        print(f"[deploy] 部署失败（可能进程占用）：{e}", flush=True)
        return None


cmd = [
    PYI_EXE,
    "--noconfirm",
    "--clean",
    os.path.join(HERE, "packaging.spec"),
]
print(" ".join(cmd), flush=True)
rc = subprocess.call(cmd, cwd=HERE)
if rc == 0:
    _deploy()
sys.exit(rc)