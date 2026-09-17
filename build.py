# Excel SQL 查询台 —— 打包脚本（onedir）
# 用法：用 venv 里的 Python 运行，例如
#   .\.venv\Scripts\python.exe build.py
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

cmd = [
    PYI_EXE,
    "--noconfirm",
    "--clean",
    os.path.join(HERE, "packaging.spec"),
]
print(" ".join(cmd), flush=True)
sys.exit(subprocess.call(cmd, cwd=HERE))