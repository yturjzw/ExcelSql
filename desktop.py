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

# QtWebEngine（pywebview gui="qt"）在远程桌面/虚拟显示（如 Todesk、RDP、虚拟机）下
# GPU 硬件加速会失效，表现为窗口白屏。经实测：仅 --disable-gpu 仍会白屏（QtWebEngine
# 6.x 仍尝试创建 GL 上下文 virgl 失败），必须强制 ANGLE 走 SwiftShader 纯软件渲染。
# 须早于任何 Qt 初始化，故放模块顶部。
os.environ["QTWEBENGINE_CHROMIUM_FLAGS"] = "--disable-gpu --use-angle=swiftshader --no-sandbox"
# Chromium sandbox 在 PyInstaller windowed(无控制台) 子系统的进程里初始化失败 → 渲染进程
# QtWebEngineProcess.exe 起不来 → 页面加载 ok=False、窗口白屏。禁用 sandbox（Qt 官方变量）。
os.environ["QTWEBENGINE_DISABLE_SANDBOX"] = "1"

# 修复白屏根因：PyInstaller 冻结后 Qt 无法靠 QLibraryInfo/qt.conf 定位到 Chromium
# 渲染进程 QtWebEngineProcess.exe（导致页面加载 ok=False、窗口灰/白屏），须显式指定
# 其在包内（_MEIPASS/PySide6/）的位置，绕过 QLibraryInfo 的路径推导。
if getattr(sys, "frozen", False):
    _proc = os.path.join(getattr(sys, "_MEIPASS", ""), "PySide6", "QtWebEngineProcess.exe")
    if os.path.exists(_proc):
        os.environ["QTWEBENGINEPROCESS_PATH"] = _proc

# 单实例互斥锁：用 Local 命名空间（按会话隔离）。Global\ 需要 SeCreateGlobalPrivilege，
# 受限制令牌下创建会失败进而误判为“已有实例”，导致应用静默退出（历史踩坑）。
MUTEX_NAME = "Local\\ExcelSqlConsole_SingleInstance"

WINDOW_TITLE = "EXCEL → SQL 查询台"
WINDOW_SIZE = (1280, 820)

# 浏览器降级模式：无任何请求超过该时长自动退出（避免后台残留进程）
_FALLBACK_IDLE_EXIT_S = 30 * 60

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

    # 验证窗口后端链路：pywebview qt（QtWebEngine / PySide6），不依赖 pythonnet。
    # 历史：edgechromium(WebView2) 需要 pythonnet，而它在 PyInstaller 冻结产物里每次都
    # 初始化失败（Failed to resolve Python.Runtime.Loader.Initialize），已弃用统一改 qt。
    try:
        import webview  # noqa: F401
        import webview.platforms.qt  # noqa: F401  触发 qtpy/PySide6 导入
        from PySide6.QtWebEngineWidgets import QWebEngineView  # noqa: F401
        emit("window backend qt: importable (PySide6/QtWebEngine OK)")
    except Exception as e:  # noqa: BLE001
        emit(f"window backend qt: IMPORT FAILED {e}")

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


def _dllprobe() -> int:
    """通过 clr_loader 原生链路逐位点尝试加载 Python.Runtime.dll（定位冻结失败根因）。

    把 Python.Runtime.dll 复制到多个目录逐个调用 pyclr_get_function，再用
    get_netfx()（无 domain ＝ 根 AppDomain，与 pythonnet 真实调用完全一致）做对照：
      · 全部位点 FAIL  → 进程/CLR 宿主层问题（与路径无关）
      · 仅原始位点 FAIL → 路径/目录相关问题
      · 新域 OK 但根域 FAIL → 根 AppDomain 特定问题
    """
    import shutil

    lines = []

    def out(s: str) -> None:
        lines.append(s)
        _log(s)

    base = getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.abspath(__file__))
    src = os.path.join(base, "pythonnet", "runtime", "Python.Runtime.dll")
    candidates = [("原始位点(_MEIPASS\\pythonnet\\runtime)", src)]

    copy_slots = [
        ("_MEIPASS 根目录", base),
        ("onedir 根目录(与 exe 同级)", os.path.dirname(base)),
        ("当前工作目录", os.getcwd()),
        ("LOCALAPPDATA\\ExcelSqlConsole", os.path.join(os.environ.get("LOCALAPPDATA", ""), "ExcelSqlConsole")),
        ("系统 TEMP", os.environ.get("TEMP", "")),
    ]
    for label, d in copy_slots:
        try:
            os.makedirs(d, exist_ok=True)
            dst = os.path.join(d, "Python.Runtime.dll")
            shutil.copyfile(src, dst)
            candidates.append((label, dst))
        except Exception as e:  # noqa: BLE001
            out(f"copy[{label}] failed: {e}")

    try:
        from clr_loader import ffi as _ffipkg

        _fw = _ffipkg.load_netfx()
        _fw.pyclr_initialize()
    except Exception as e:  # noqa: BLE001
        out(f"load_netfx/pyclr_initialize failed: {type(e).__name__}: {e}")
        _write_diagnose("\n".join(lines))
        return 1

    for label, path in candidates:
        try:
            t0 = time.monotonic()
            _dom = _fw.pyclr_create_appdomain(b"dllprobe", _ffipkg.ffi.NULL)
            fn = _fw.pyclr_get_function(
                _dom, path.encode("utf-8"), b"Python.Runtime.Loader", b"Initialize"
            )
            dt = time.monotonic() - t0
            verdict = "OK  " if fn != _ffipkg.ffi.NULL else "FAIL"
            out(f"{verdict} {label}: {path} ({dt * 1000:.0f} ms)")
        except Exception as e:  # noqa: BLE001
            out(f"ERR  {label}: {path}: {type(e).__name__}: {e}")

    # 对照：完全复刻 pythonnet 的真实调用（get_netfx 无 domain = 根 AppDomain）
    try:
        from clr_loader import get_netfx

        rt = get_netfx()
        asm = rt.get_assembly(src)
        f = asm.get_function("Python.Runtime.Loader.Initialize")
        out(f"get_netfx(根域)原始位点 -> OK: {f!r}")
    except Exception as e:  # noqa: BLE001
        out(f"get_netfx(根域)原始位点 -> FAIL: {type(e).__name__}: {e}")

    out("== dllprobe 完成 ==")
    _write_diagnose("\n".join(lines))
    return 0


def _rendertest() -> int:
    """冻结环境 QtWebEngine 渲染自检：离屏加载红底页并截图判断是否白屏。

    定位「源码能渲染、冻结白屏」的差异：打印实际生效的渲染环境变量，再抓离屏
    像素判断软件渲染是否真的在冻结产物里工作。结果写 diagnose 文件与 rendertest.png。
    """
    lines = []

    def out(s):
        lines.append(s)
        _log(s)

    out(f"QTWEBENGINE_CHROMIUM_FLAGS={os.environ.get('QTWEBENGINE_CHROMIUM_FLAGS')!r}")
    out(f"QT_OPENGL={os.environ.get('QT_OPENGL')!r}")
    out(f"QT_ANGLE_PLATFORM={os.environ.get('QT_ANGLE_PLATFORM')!r}")
    out(f"QT_PLUGIN_PATH={os.environ.get('QT_PLUGIN_PATH')!r}")
    out(f"QTWEBENGINEPROCESS_PATH={os.environ.get('QTWEBENGINEPROCESS_PATH')!r}")

    png = _exe_dir_path("rendertest.png")
    try:
        from PySide6.QtWidgets import QApplication
        from PySide6.QtWebEngineWidgets import QWebEngineView
        from PySide6.QtCore import QTimer, QEventLoop

        app = QApplication([])
        view = QWebEngineView()
        view.setHtml(
            "<html><body style='margin:0;background:#ff0000'>"
            "<h1 style='color:#fff;font-size:48px'>RT</h1></body></html>"
        )
        view.resize(600, 400)
        view.show()
        state = {"ok": None}
        loop = QEventLoop()

        def on_finished(ok):
            state["ok"] = ok
            loop.quit()

        view.loadFinished.connect(on_finished)
        QTimer.singleShot(12000, loop.quit)  # 兜底超时
        loop.exec()
        app.processEvents()

        pm = view.grab()
        pm.save(png)
        img = pm.toImage()
        cx, cy = img.width() // 2, img.height() // 2
        c = img.pixelColor(cx, cy)
        nonwhite = c.red() != 255 or c.green() != 255 or c.blue() != 255
        out(
            f"load_finished ok={state['ok']} "
            f"rendered={pm.width()}x{pm.height()} center RGB=({c.red()},{c.green()},{c.blue()}) "
            f"-> {'渲染OK' if nonwhite else '白屏'}"
        )
        app.quit()
    except Exception as e:  # noqa: BLE001
        out(f"rendertest FAILED: {type(e).__name__}: {e}")
        out(traceback.format_exc())

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


def _app_icon_path() -> str:
    """窗口图标 assets/app.ico 路径：冻结后 _MEIPASS，源码为项目目录。"""
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", None) or os.path.dirname(sys.executable)
        return os.path.join(base, "assets", "app.ico")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "app.ico")


def _ensure_native_dll_paths() -> None:
    """把冻结包内的原生 DLL 目录加入进程的 DLL 搜索路径。

    背景（pythonnet + PyInstaller 已知坑）：clr_loader 通过 LoadLibrary 加载
    Python.Runtime.dll，而它导入表依赖的 python312.dll 等原生库位于 onedir 的
    _internal（sys._MEIPASS）根目录，并不在系统 DLL 搜索路径上。不显式加入时，
    LoadLibrary 返回 NULL，表现为初始化报
    'Failed to resolve Python.Runtime.Loader.Initialize from ...Python.Runtime.dll'。
    源码运行加载路径天然命中，无需此处理。
    """
    base = getattr(sys, "_MEIPASS", None)
    if not base:
        return
    candidates = [base]
    for sub in ("pythonnet", os.path.join("pythonnet", "runtime"),
                "clr_loader", os.path.join("clr_loader", "ffi", "dlls", "amd64")):
        candidates.append(os.path.join(base, sub))
    for p in candidates:
        if os.path.isdir(p):
            try:
                os.add_dll_directory(p)
            except (AttributeError, OSError):
                pass
    # PATH 也会被无标志的 LoadLibraryW 搜索到；ClrLoader（原生 C++）加载
    # Python.Runtime.dll 时并不一定走 add_dll_directory 目录，双保险。
    try:
        os.environ["PATH"] = base + os.pathsep + os.environ.get("PATH", "")
    except Exception:  # noqa: BLE001
        pass


def main() -> int:
    # 无窗口诊断模式（打包产物自检用）
    if "--diagnose" in sys.argv:
        return _diagnose()
    # 无窗口 DLL 加载位点探测（定位 pythonnet/python.Runtime.dll 冻结加载失败根因用）
    if "--dllprobe" in sys.argv:
        return _dllprobe()
    # 无窗口 QtWebEngine 渲染自检（离屏加载红底页截图判色，定位白屏根因用）
    if "--rendertest" in sys.argv:
        return _rendertest()

    # 单实例：已有实例时本进程直接退出
    mutex = _acquire_single_instance()
    if mutex is None:
        # 已有实例在运行：把它的查询台页面重新在系统浏览器中打开（避免“双击没反应”）
        return _reopen_existing()

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
    _write_last_port(port)  # 供「已有实例时重新打开浏览器页面」使用
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

    # 3) 弹出原生窗口（QtWebEngine 内嵌 Chromium，不依赖 pythonnet/.NET 桥）。
    #    历史：edgechromium(WebView2) 需要 pythonnet，而它在 PyInstaller 冻结产物里
    #    每次都初始化失败（Failed to resolve Python.Runtime.Loader.Initialize），
    #    已弃用，统一改用 Qt 后端；Qt 也不可用时才降级为系统浏览器。
    try:
        try:
            import webview
        except Exception as e:  # noqa: BLE001
            raise RuntimeError(f"无法加载窗口组件（QtWebEngine）：{e}") from e

        window = webview.create_window(
            WINDOW_TITLE,
            url,
            width=WINDOW_SIZE[0],
            height=WINDOW_SIZE[1],
            min_size=(960, 640),
            confirm_close=False,
        )
        webview.start(gui="qt", icon=_app_icon_path())  # QtWebEngine 原生窗口；阻塞至窗口关闭
    except Exception as e:  # noqa: BLE001
        # 窗口链路失败时不退出，降级为「系统浏览器打开内嵌服务」，保证产物仍可用。
        _log(f"desktop window unavailable -> browser fallback:\n{traceback.format_exc()}")
        exit_now = _fallback_to_browser(url, str(e))
        if not exit_now:
            # 用户选择「保持运行」：后台服务继续，直到空闲超时或服务线程退出。
            _wait_fallback_idle(backend_app, server_thread)
    finally:
        # 4) 窗口关闭（或降级后）优雅停服务
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


def _message_box(title: str, text: str, *, icon_error: bool = True) -> None:
    """信息/错误提示框（不阻塞后端服务线程）。"""
    try:
        flags = 0x10 if icon_error else 0x40  # MB_ICONERROR | MB_ICONINFORMATION
        ctypes.windll.user32.MessageBoxW(0, text, title, flags)
    except Exception:  # noqa: BLE001
        try:
            print(f"{title}: {text}", file=sys.stderr)
        except Exception:  # noqa: BLE001
            pass


def _message_box_yesno(title: str, text: str, *, default_no: bool = True) -> bool:
    """问询对话框：返回 True=选「是（退出应用）」，False=选「否（保持运行）」。"""
    flags = 0x04 | 0x40  # MB_YESNO | MB_ICONINFORMATION
    if default_no:
        flags |= 0x00000100  # MB_DEFBUTTON2：默认选中「否」，回车不会误退出
    try:
        rc = ctypes.windll.user32.MessageBoxW(0, text, title, flags)
        return rc == 6  # IDYES
    except Exception:  # noqa: BLE001
        try:
            print(f"{title}: {text}", file=sys.stderr)
        except Exception:  # noqa: BLE001
            pass
        return True  # 弹不出对话框时按「退出」处理，避免留僵尸进程


def _fallback_to_browser(url: str, _reason: str) -> bool:
    """窗口链路失败时的浏览器降级：打开系统浏览器访问内嵌服务。

    返回 True＝用户选择立即退出应用；False＝保持后台运行（继续服务）。
    失败的细节只进日志，不再把「失败」字样挂在用户面前。

    由于冻结产物的原生窗口链路每次都会失败（见 _diagnose 的 pythonnet 定位结论），
    这条说明只在**首次进入浏览器模式**时弹出一次；用户选「保持运行」后写入标记，
    之后每次启动直接静默打开浏览器，不再打扰。
    """
    try:
        import webbrowser

        if not webbrowser.open(url, new=1):
            _log("webbrowser.open returned False")
    except Exception as e:  # noqa: BLE001
        _log(f"webbrowser.open failed: {e}")

    # 已记住「保持运行」：静默降级，不再弹提示。
    if _fallback_silent_remembered():
        _log("browser fallback: silent (remembered keep-running)")
        return False

    exit_now = _message_box_yesno(
        "Excel SQL 查询台",
        "已自动改用系统浏览器打开查询台。\n"
        "页面功能与桌面窗口完全一致（查询 Excel / Doris / Oracle 等）。\n\n"
        "· 「退出应用」＝结束后台服务，浏览器页面也会随之失效；\n"
        "· 「保持运行」＝关闭本提示，查询台继续在浏览器中工作。\n\n"
        "此提示只在首次出现一次，选择「保持运行」后本次内不再重复；\n"
        f"查询台空闲 {_FALLBACK_IDLE_EXIT_S // 60} 分钟会自动退出，\n"
        "使用中再次双击 ExcelSqlConsole 可重新打开页面。",
        default_no=True,
    )
    if not exit_now:
        _remember_fallback_silent()
    return exit_now


def _wait_fallback_idle(backend_app, server_thread) -> None:
    """浏览器降级模式的保活循环：空闲超时或服务线程退出后返回（停服由调用方 finally 完成）。

    降级模式是最后手段，任何小异常都不能让应用「致命错误」退出：读时间戳失败时
    跳过空闲判定（按永不空闲处理），只靠服务线程存活状态兜底。
    """
    while server_thread.is_alive():
        time.sleep(5)
        try:
            last = getattr(backend_app, "last_request_at", None)
            if last is not None and (time.time() - last) >= _FALLBACK_IDLE_EXIT_S:
                _log(f"browser fallback idle timeout ({_FALLBACK_IDLE_EXIT_S}s) -> exit")
                break
        except Exception:  # noqa: BLE001
            _log(f"browser fallback idle check failed:\n{traceback.format_exc()}")
    if not server_thread.is_alive():
        _log("browser fallback: backend server thread exited unexpectedly -> shutting down")


def _last_port_path() -> str:
    """last_port.txt 路径：LOCALAPPDATA；不可用则回退 exe 所在目录。"""
    local = os.environ.get("LOCALAPPDATA")
    if local:
        try:
            d = os.path.join(local, "ExcelSqlConsole")
            os.makedirs(d, exist_ok=True)
            return os.path.join(d, "last_port.txt")
        except OSError:
            pass
    return _exe_dir_path("last_port.txt")


def _write_last_port(port: int) -> None:
    """记录本次服务的端口（供「已有实例时重新打开页面」使用）。

    LOCALAPPDATA 与 exe 目录双路径尝试，避免某个位置不可写时静默丢失。
    """
    targets = [_last_port_path(), _exe_dir_path("last_port.txt")]
    seen = set()
    for t in targets:
        if t in seen:
            continue
        seen.add(t)
        try:
            os.makedirs(os.path.dirname(t), exist_ok=True)
            with open(t, "w", encoding="utf-8") as f:
                f.write(str(port))
            return
        except Exception:  # noqa: BLE001
            continue


def _read_last_port() -> int | None:
    try:
        with open(_last_port_path(), "r", encoding="utf-8") as f:
            return int(f.read().strip())
    except Exception:  # noqa: BLE001
        return None


def _fallback_flag_path() -> str:
    """fallback_silent.flag 路径：LOCALAPPDATA；不可用则回退 exe 所在目录。"""
    local = os.environ.get("LOCALAPPDATA")
    if local:
        try:
            d = os.path.join(local, "ExcelSqlConsole")
            os.makedirs(d, exist_ok=True)
            return os.path.join(d, "fallback_silent.flag")
        except OSError:
            pass
    return _exe_dir_path("fallback_silent.flag")


def _fallback_silent_remembered() -> bool:
    """用户是否已选择过「保持运行」（此后浏览器降级静默进行）。"""
    return os.path.exists(_fallback_flag_path())


def _remember_fallback_silent() -> None:
    """记录「保持运行」选择（此后浏览器降级不再弹提示）。"""
    try:
        p = _fallback_flag_path()
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            f.write("keep_running")
    except Exception:  # noqa: BLE001
        _log(f"write fallback flag failed:\n{traceback.format_exc()}")


def _is_url_alive(url: str, timeout: float = 1.0) -> bool:
    import http.client
    import urllib.parse

    try:
        host, port = urllib.parse.urlparse(url).netloc.split(":")
        conn = http.client.HTTPConnection(host, int(port), timeout=timeout)
        conn.request("GET", "/api/health")
        ok = conn.getresponse().status == 200
        conn.close()
        return ok
    except Exception:  # noqa: BLE001
        return False


def _reopen_existing() -> int:
    """单实例已被占用：把已有实例的查询台页面重新在系统浏览器中打开。"""
    port = _read_last_port()
    url = f"http://127.0.0.1:{port}/" if port else None
    if url and _is_url_alive(url):
        try:
            import webbrowser

            webbrowser.open(url, new=1)
            _log(f"existing instance on port {port} -> reopened in browser")
        except Exception:  # noqa: BLE001
            pass
    else:
        _log(f"single-instance lock busy but no live server (port={port}) -> exiting")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception:  # noqa: BLE001
        _log(f"fatal:\n{traceback.format_exc()}")
        _message_box("致命错误", traceback.format_exc())
        raise SystemExit(1)