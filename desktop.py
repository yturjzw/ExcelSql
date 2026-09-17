"""Excel SQL 查询台 —— Windows 桌面启动器。

职责：
  1. 单实例互斥锁：再次双击只唤醒已有窗口，不重复起服务。
  2. 选一个空闲端口，在后台线程里启动 FastAPI（uvicorn）。
  3. 等后端就绪后弹出 pywebview 原生窗口（走系统 WebView2），
     无浏览器地址栏，形同真正的桌面应用。
  4. 窗口关闭时优雅停止服务（shutdown 信号 + uvicorn 停止 + 线程回收）。

打包（PyInstaller onedir）后作为入口：桌面窗口 + 内嵌后端一体，离线可用。
"""
from __future__ import annotations

import ctypes
import os
import sys
import threading
import time
import traceback

MUTEX_NAME = "Global\\ExcelSqlConsole_SingleInstance"

WINDOW_TITLE = "EXCEL → SQL 查询台"
WINDOW_SIZE = (1280, 820)

# ---- P0 内存优化：给 WebView2 追加降内存的 Chromium 启动参数 ----
# 本应用是纯静态表格页面，无视频/WebGL/游戏，不需要 GPU 加速；
# 关掉 GPU 进程 + 限制渲染进程数，能显著降低常驻内存（实测基线的
# gpu-process ~75MB / 多 utility 进程）。
#
# 注意：pywebview 6.2.1 用 CoreWebView2CreationProperties.AdditionalBrowserArguments
# 硬编码了浏览器参数，该属性会覆盖 WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS 环境变量。
# 因此这些参数必须通过 monkey-patch 注入到该属性，而不是写环境变量。
# 已实测确认参数成功进入浏览器进程命令行（--disable-gpu 等可见）。
WEBVIEW2_MEMORY_ARGS = " ".join([
    "--disable-gpu",                       # 禁用硬件加速（避免 GPU 驱动兼容问题，静态表格不需要）
    "--renderer-process-limit=1",          # 限制渲染进程数为 1（本应用本就单页面）
    "--disable-features=CalculateNativeWinOcclusion",  # 省掉窗口遮挡计算
    "--no-first-run",                      # 跳过首次运行初始化
])

_EDGE_PATCHED = False


def _patch_edgechromium_args() -> None:
    """monkey-patch pywebview 的 EdgeChrome.__init__，注入降内存参数。

    pywebview 6.2.1 的 edgechromium.EdgeChrome.__init__ 把浏览器参数硬编码到
    CoreWebView2CreationProperties.AdditionalBrowserArguments（该属性优先级高于
    WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS 环境变量）。这里在调用 webview.start()
    之前替换掉 __init__，在原参数基础上追加 WEBVIEW2_MEMORY_ARGS。
    """
    global _EDGE_PATCHED
    if _EDGE_PATCHED:
        return

    import webview.platforms.edgechromium as ec

    def _patched_init(self, form, window, cache_dir):
        self.pywebview_window = window
        self.webview = ec.WebView2()
        props = ec.CoreWebView2CreationProperties()

        runtime_path = ec.webview_settings["WEBVIEW2_RUNTIME_PATH"]
        if runtime_path:
            if not ec.os.path.isabs(runtime_path):
                runtime_path = ec.os.path.join(ec.get_app_root(), runtime_path)
            if ec.os.path.exists(runtime_path):
                props.BrowserExecutableFolder = runtime_path
                ec.logger.debug(f"Using custom WebView2 runtime: {runtime_path}")
            else:
                ec.logger.warning(
                    f"Custom WebView2 runtime path does not exist: {runtime_path}. Using system WebView2."
                )

        props.UserDataFolder = cache_dir
        self.user_data_folder = props.UserDataFolder
        props.set_IsInPrivateModeEnabled(ec._state["private_mode"])
        # 追加降内存参数（原硬编码为 '--disable-features=ElasticOverscroll'）
        props.AdditionalBrowserArguments = (
            "--disable-features=ElasticOverscroll " + WEBVIEW2_MEMORY_ARGS
        )

        if ec.webview_settings["ALLOW_FILE_URLS"]:
            props.AdditionalBrowserArguments += " --allow-file-access-from-files"

        if ec.webview_settings["REMOTE_DEBUGGING_PORT"] is not None:
            props.AdditionalBrowserArguments += (
                f' --remote-debugging-port={ec.webview_settings["REMOTE_DEBUGGING_PORT"]}'
            )

        self.webview.CreationProperties = props

        self.form = form
        form.Controls.Add(self.webview)

        self.js_results = {}
        self.js_result_semaphore = ec.Semaphore(0)
        self.webview.Dock = ec.WinForms.DockStyle.Fill
        self.webview.BringToFront()
        self.webview.CoreWebView2InitializationCompleted += self.on_webview_ready
        self.webview.NavigationStarting += self.on_navigation_start
        self.webview.NavigationCompleted += self.on_navigation_completed
        self.webview.WebMessageReceived += self.on_script_notify
        self.syncContextTaskScheduler = ec.TaskScheduler.FromCurrentSynchronizationContext()
        self.webview.DefaultBackgroundColor = ec.Color.FromArgb(
            255,
            int(window.background_color.lstrip("#")[0:2], 16),
            int(window.background_color.lstrip("#")[2:4], 16),
            int(window.background_color.lstrip("#")[4:6], 16),
        )

        if window.transparent:
            self.webview.DefaultBackgroundColor = ec.Color.Transparent

        self.url = None
        self.ishtml = False
        self.html = ec.DEFAULT_HTML

        self.webview.EnsureCoreWebView2Async(None)

    ec.EdgeChrome.__init__ = _patched_init
    _EDGE_PATCHED = True

LOG_PATH = os.path.join(
    os.path.expandvars(r"%LOCALAPPDATA%"),
    "ExcelSqlConsole",
    "app.log",
)

DIAG_PATH = os.path.join(
    os.path.expandvars(r"%LOCALAPPDATA%"),
    "ExcelSqlConsole",
    "diagnose.txt",
)


def _log(msg: str) -> None:
    """把诊断信息追加到日志文件（windowed exe 无控制台，stdout/stderr 常为 None）。"""
    for target in (LOG_PATH, _exe_dir_path("app.log")):
        try:
            os.makedirs(os.path.dirname(target), exist_ok=True)
            with open(target, "a", encoding="utf-8") as f:
                f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n")
            return
        except Exception:  # noqa: BLE001
            continue


def _acquire_single_instance() -> object | None:
    """返回一个 Windows 命名互斥锁句柄；已被占用则返回 None（并唤起已有窗口）。"""
    import ctypes
    from ctypes import wintypes

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
    kernel32.CreateMutexW.restype = wintypes.HANDLE
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    kernel32.CloseHandle.restype = wintypes.BOOL

    handle = kernel32.CreateMutexW(None, False, MUTEX_NAME)
    if not handle:
        return None
    # ERROR_ALREADY_EXISTS (183)：已有实例在运行
    if ctypes.get_last_error() == 183:
        kernel32.CloseHandle(handle)
        _bring_to_front()
        return None
    return handle


def _bring_to_front() -> None:
    """尽力唤醒已有实例的窗口（按窗口标题查找并置顶）。"""
    import ctypes
    from ctypes import wintypes

    try:
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        user32.FindWindowW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR]
        user32.FindWindowW.restype = wintypes.HWND
        user32.SetForegroundWindow.argtypes = [wintypes.HWND]
        user32.SetForegroundWindow.restype = wintypes.BOOL
        hwnd = user32.FindWindowW(None, WINDOW_TITLE)
        if hwnd:
            user32.SetForegroundWindow(hwnd)
    except Exception:
        # 置顶失败不影响「单实例」语义，静默即可
        pass


def _server_target(backend_app, host: str, port: int) -> None:
    """服务线程入口：捕获任何异常并写日志，避免静默退出。"""
    try:
        backend_app.run_server(host=host, port=port)
    except Exception:  # noqa: BLE001
        _log(f"server thread crashed:\n{traceback.format_exc()}")


def _diagnose() -> int:
    """无窗口诊断模式：仅启动后端，验证离线加载 excel 扩展与前端，结果写入文件后退出。

    用法：ExcelSqlConsole.exe --diagnose
    结果写入 %LOCALAPPDATA%\\ExcelSqlConsole\\diagnose.txt（windowed exe 无控制台）。
    """
    lines: list[str] = []

    def emit(msg: str) -> None:
        lines.append(msg)
        try:
            # console 构建时有 stdout，便于直接观察；windowed 构建 stdout 为 None
            if sys.stdout is not None:
                sys.stdout.write(msg + "\n")
                sys.stdout.flush()
        except Exception:  # noqa: BLE001
            pass

    import duckdb
    from backend import app as backend_app

    emit("== ExcelSqlConsole 诊断 ==")
    emit(f"frozen: {bool(getattr(sys, 'frozen', False))}")
    emit(f"frontend exists: {backend_app.FRONTEND_DIR.exists()} ({backend_app.FRONTEND_DIR})")

    con = duckdb.connect()
    try:
        backend_app._load_excel(con)
        row = con.execute(
            "SELECT extension_name, loaded FROM duckdb_extensions() WHERE extension_name='excel'"
        ).fetchone()
        emit(f"excel extension: {row}")
    except Exception as e:  # noqa: BLE001
        emit(f"excel extension FAILED: {e}")
        _write_diagnose("\n".join(lines))
        return 1
    finally:
        con.close()

    # 验证 openpyxl（xlsx 导出依赖）在冻结环境下可用
    try:
        import openpyxl  # noqa: F401
        emit("openpyxl: importable")
    except Exception as e:  # noqa: BLE001
        emit(f"openpyxl: IMPORT FAILED {e}")

    # 起服务并访问 /api/health 与 /，验证完整链路
    port = backend_app.find_free_port()
    t = threading.Thread(
        target=_server_target,
        kwargs={"backend_app": backend_app, "host": "127.0.0.1", "port": port},
        name="uvicorn",
        daemon=True,
    )
    t.start()
    url = f"http://127.0.0.1:{port}/"
    if not _wait_until_ready(url, timeout=15.0):
        emit("backend not ready")
        emit(f"server thread alive: {t.is_alive()}")
        backend_app.shutdown()
        _write_diagnose("\n".join(lines))
        return 1
    import http.client
    import urllib.parse

    hp = urllib.parse.urlparse(url).netloc.split(":")
    conn = http.client.HTTPConnection(hp[0], int(hp[1]), timeout=2.0)
    conn.request("GET", "/")
    body = conn.getresponse().read().decode("utf-8", "replace")
    conn.close()
    emit(f"index.html loaded chars: {len(body)}")
    emit(f"has SELECT placeholder: {'SELECT * FROM data' in body}")
    backend_app.shutdown()
    t.join(timeout=3)
    emit("== 诊断完成 ==")
    _write_diagnose("\n".join(lines))
    return 0


def _write_diagnose(text: str) -> None:
    """把诊断结果写到 LOCALAPPDATA；若失败，回退写到 exe 所在目录（便于排查/沙箱内测试）。"""
    for target in (DIAG_PATH, _exe_dir_path("diagnose.txt")):
        try:
            os.makedirs(os.path.dirname(target), exist_ok=True)
            with open(target, "w", encoding="utf-8") as f:
                f.write(text + "\n")
            return
        except Exception:  # noqa: BLE001
            continue


def _exe_dir_path(name: str) -> str:
    """返回 exe 所在目录下的一个文件路径（冻结时 = dist 目录，源码时 = 项目目录）。"""
    if getattr(sys, "frozen", False):
        return os.path.join(os.path.dirname(sys.executable), name)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), name)


def main() -> int:
    # 无窗口诊断模式（打包产物自检用）
    if "--diagnose" in sys.argv:
        return _diagnose()

    # 单实例：已有实例时本进程直接退出
    mutex = _acquire_single_instance()
    if mutex is None:
        return 0

    backend = None
    server_thread = None
    try:
        # 延迟导入：PyInstaller 冷启动更快，且此处才需要完整后端
        from backend import app as backend_app

        backend = backend_app
    except Exception as e:  # noqa: BLE001
        _log(f"backend import failed:\n{traceback.format_exc()}")
        _message_box("启动失败", f"无法加载后端模块：\n{e}")
        return 1

    # 1) 选空闲端口并后台启动 uvicorn
    port = backend_app.find_free_port()
    server_thread = threading.Thread(
        target=_server_target,
        kwargs={"backend_app": backend_app, "host": "127.0.0.1", "port": port},
        name="uvicorn",
        daemon=True,
    )
    server_thread.start()

    # 2) 等后端就绪（最多 15 秒）
    url = f"http://127.0.0.1:{port}/"
    ready = _wait_until_ready(url, timeout=15.0)
    if not ready:
        _log(f"backend not ready within timeout (thread alive={server_thread.is_alive()})")
        backend_app.shutdown()
        _message_box("启动失败", "后端服务未能就绪，请稍后重试。")
        return 1

    # 3) 弹出原生窗口
    try:
        import webview
    except Exception as e:  # noqa: BLE001
        _log(f"webview import failed:\n{traceback.format_exc()}")
        backend_app.shutdown()
        _message_box("启动失败", f"无法加载窗口组件（WebView2）：\n{e}")
        return 1

    # P0 内存优化：必须在 webview.start()（触发 WebView2 实例化）前 patch
    try:
        _patch_edgechromium_args()
    except Exception as e:  # noqa: BLE001
        _log(f"edgechromium patch failed (fallback to default args):\n{traceback.format_exc()}")

    try:
        window = webview.create_window(
            WINDOW_TITLE,
            url,
            width=WINDOW_SIZE[0],
            height=WINDOW_SIZE[1],
            min_size=(960, 640),
            confirm_close=False,
        )
    except Exception as e:  # noqa: BLE001
        _log(f"create_window failed:\n{traceback.format_exc()}")
        backend_app.shutdown()
        _message_box("启动失败", f"创建窗口失败：\n{e}")
        return 1

    try:
        webview.start(gui="edgechromium")  # Windows 下走 WebView2
    except Exception as e:  # noqa: BLE001
        _log(f"webview.start failed:\n{traceback.format_exc()}")
        _message_box("运行出错", f"窗口运行失败：\n{e}")
    finally:
        # 4) 窗口关闭：优雅停服务
        backend_app.shutdown()
        if server_thread is not None:
            server_thread.join(timeout=3)

    if mutex is not None:
        try:
            ctypes.windll.kernel32.CloseHandle(mutex)
        except Exception:  # noqa: BLE001
            pass
    return 0


def _wait_until_ready(url: str, timeout: float) -> bool:
    import http.client
    import urllib.parse

    deadline = time.monotonic() + timeout
    host, port = urllib.parse.urlparse(url).netloc.split(":")
    while time.monotonic() < deadline:
        try:
            conn = http.client.HTTPConnection(host, int(port), timeout=1.0)
            conn.request("GET", "/api/health")
            resp = conn.getresponse()
            conn.close()
            if resp.status == 200:
                return True
        except Exception:  # noqa: BLE001
            time.sleep(0.15)
    return False


def _message_box(title: str, text: str) -> None:
    try:
        ctypes.windll.user32.MessageBoxW(0, text, title, 0x10)  # MB_ICONERROR
    except Exception:  # noqa: BLE001
        try:
            print(f"{title}: {text}", file=sys.stderr)
        except Exception:  # noqa: BLE001
            pass


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception:  # noqa: BLE001
        _log(f"fatal:\n{traceback.format_exc()}")
        _message_box("致命错误", traceback.format_exc())
        raise SystemExit(1)