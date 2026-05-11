import threading
import time
import os

# Ensure WebView2 receives this browser argument before pywebview initializes.
# This disables the MS WebBrowser drag feature which can allow frameless
# WebView2 windows to be moved by dragging the web content.
os.environ.setdefault("WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS", "--disable-features=msWebBrowserDrag")

import webview
import app
from pathlib import Path
import json
import requests
import os
import tempfile
import subprocess
import sys
import traceback
import ctypes

# -------------------------------------------------------------------
# LOGGING
# -------------------------------------------------------------------
def _log_path():
    local = os.environ.get("LOCALAPPDATA", "")
    log_dir = Path(local) / "BrewInsPOS"
    log_dir.mkdir(parents=True, exist_ok=True)
    return log_dir / "pos_window.log"

LOG_FILE = _log_path()

def log(msg: str):
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except:
        # last resort, don't crash logger
        pass

log("======================================================")
log("app_window.py starting up")

# -------------------------------------------------------------------
# Versioning
# -------------------------------------------------------------------
VERSION_FILE = Path(__file__).resolve().parent / "version.txt"

def get_installed_version():
    try:
        v = VERSION_FILE.read_text().strip()
        log(f"[VERSION] Installed version: {v}")
        return v
    except Exception as e:
        log(f"[VERSION] Failed to read version.txt: {e}")
        return "0.0.0"


VERSION_CHECK_URL = "https://raw.githubusercontent.com/https-perks/brewin-pos/main/version.json"

def check_for_updates():
    log("[UPDATE] Checking for updates...")
    try:
        response = requests.get(VERSION_CHECK_URL, timeout=5)
        update_data = response.json()
        latest = update_data.get("latest_version", "0.0.0")
        installed = get_installed_version()
        log(f"[UPDATE] Installed={installed}, Latest={latest}")

        if latest != installed:
            url = update_data.get("download_url")
            log(f"[UPDATE] Update available. URL={url}")
            return url, latest
        else:
            log("[UPDATE] No update available.")
    except Exception as e:
        log(f"[UPDATE] Update check failed: {e}\n{traceback.format_exc()}")

    return None, None

# -------------------------------------------------------------------
# Flask Startup
# -------------------------------------------------------------------
HOST = "127.0.0.1"
PORT = 5000


def run_flask():
    try:
        log("[FLASK] Starting Flask server thread...")
        app.app.run(host=HOST, port=PORT, debug=False)
        log("[FLASK] Flask server exited normally.")
    except Exception as e:
        log(f"[FLASK] Flask crashed: {e}\n{traceback.format_exc()}")

def wait_for_flask():
    """Ensure Flask is actually responding before showing POS window."""
    log("[FLASK] Waiting for Flask to respond on /splash...")
    for i in range(40):  # up to ~4 seconds
        try:
            r = requests.get(f"http://{HOST}:{PORT}/splash", timeout=0.25)
            log(f"[FLASK] Health check attempt {i+1}, status={r.status_code}")
            if r.status_code == 200:
                log("[FLASK] Flask responded OK on /splash.")
                return True
        except Exception as e:
            time.sleep(0.1)
    log("[FLASK] ERROR: Flask did not respond in time.")
    return False

# -------------------------------------------------------------------
# Update Installation Logic
# -------------------------------------------------------------------
class ApiBridge:
    def request_exit(self):
        log("[API] request_exit() called from JS")
        try:
            return confirm_exit(webview.windows[0])
        except Exception as e:
            log(f"[API] request_exit error: {e}\n{traceback.format_exc()}")
            return False

    def check_update(self):
        log("[API] check_update() called from JS")
        url, version = check_for_updates()
        return {"url": url, "version": version}

    def install_update(self, url, version):
        log(f"[API] install_update() requested. url={url}, version={version}")
        window = webview.windows[0]

        try:
            window.evaluate_js("""
                document.body.innerHTML =
                "<h1 style='font-size:40px; text-align:center; margin-top:150px;'>Updating...<br>Please wait.</h1>";
            """)
        except Exception as e:
            log(f"[API] Failed to update UI for update: {e}\n{traceback.format_exc()}")

        temp_path = os.path.join(tempfile.gettempdir(), f"BrewInsPOS_{version}.exe")
        log(f"[API] Downloading update to temp file: {temp_path}")

        try:
            r = requests.get(url, stream=True)
            with open(temp_path, "wb") as f:
                for chunk in r.iter_content(1024 * 256):
                    f.write(chunk)
            log("[API] Update download complete.")
        except Exception as e:
            log(f"[API] Update download error: {e}\n{traceback.format_exc()}")
            try:
                window.evaluate_js("alert('Update failed: download error');")
            except:
                pass
            return

        current_exe = sys.executable
        updater_exe = os.path.join(os.path.dirname(current_exe), "updater.exe")
        log(f"[API] current_exe={current_exe}")
        log(f"[API] updater_exe={updater_exe}")

        try:
            subprocess.Popen([updater_exe, current_exe, temp_path, version])
            log("[API] Launched updater.exe, exiting POS app.")
        except Exception as e:
            log(f"[API] Updater launch failed: {e}\n{traceback.format_exc()}")

        try:
            window.destroy()
        except:
            pass

        os._exit(0)

# -------------------------------------------------------------------
# Settings and Security PIN Validation
# -------------------------------------------------------------------
def get_pos_settings():
    """Fetch POS settings from Flask."""
    try:
        r = requests.get(f"http://{HOST}:{PORT}/api/admin/settings", timeout=1)
        if r.status_code == 200:
            return r.json()
    except Exception as e:
        log(f"[EXIT] Failed to fetch settings: {e}")
    return {}


def validate_security_pin(pin: str) -> bool:
    """Validate entered PIN using Flask/security API."""
    try:
        r = requests.post(
            f"http://{HOST}:{PORT}/api/security/validate_pin",
            json={"pin": pin},
            timeout=1
        )

        if r.status_code == 200:
            data = r.json()
            return bool(data.get("ok"))

    except Exception as e:
        log(f"[EXIT] Failed to validate security PIN: {e}")

    return False


def has_admin_pins() -> bool:
    """Return True if at least one admin PIN exists."""
    try:
        r = requests.get(f"http://{HOST}:{PORT}/api/admin/admin_pins", timeout=1)
        if r.status_code == 200:
            pins = r.json()
            return len(pins) > 0
    except Exception as e:
        log(f"[EXIT] Failed to check admin pins: {e}")

    return False


def get_active_shift():
    """Return active shift info from Flask, or empty dict if none."""
    try:
        r = requests.get(f"http://{HOST}:{PORT}/api/shift/active", timeout=1)
        if r.status_code == 200:
            return r.json() or {}
    except Exception as e:
        log(f"[EXIT] Failed to check active shift: {e}")

    return {}

# -------------------------------------------------------------------
# Exit Logic
# -------------------------------------------------------------------
def confirm_exit(window):
    log("[EXIT] confirm_exit() called.")

    # ------------------------------------------------------------
    # 1) Active-shift safety check
    # ------------------------------------------------------------
    active_shift = get_active_shift()

    if active_shift.get("id"):
        cashier = active_shift.get("cashier", "Unknown")
        started = active_shift.get("ts_start", "unknown time")

        log(
            f"[EXIT] Active shift detected: "
            f"cashier={cashier}, started={started}"
        )

        # Ask whether this is an intentional force-exit.
        try:
            force_exit = window.evaluate_js(f"""
                confirm(
                    "A shift is still active.\\n\\n" +
                    "Cashier: {cashier}\\n" +
                    "Started: {started}\\n\\n" +
                    "You should close the shift before exiting.\\n\\n" +
                    "Force exit anyway? The shift will remain open."
                );
            """)
        except Exception as e:
            log(f"[EXIT] Failed to show force-exit warning: {e}")
            return False

        if not force_exit:
            log("[EXIT] User cancelled exit because shift is active.")
            return False

        # If there are no admin PINs, do NOT allow force exit.
        # This prevents a student from bypassing the warning if no PIN exists.
        if not has_admin_pins():
            log("[EXIT] Force exit blocked: no admin PINs exist.")
            try:
                window.evaluate_js("""
                    alert(
                        "Force exit is not available because no Admin PINs exist.\\n\\n" +
                        "Please close the shift before exiting."
                    );
                """)
            except:
                pass
            return False

        # Require admin PIN for force exit.
        try:
            pin = window.evaluate_js("prompt('Enter Admin PIN to force exit:')")
        except Exception as e:
            log(f"[EXIT] Error prompting for force-exit PIN: {e}")
            return False

        if pin is None:
            log("[EXIT] Force-exit PIN prompt cancelled.")
            return False

        log("[EXIT] Force-exit PIN entered; validating...")

        if not validate_security_pin(pin):
            log("[EXIT] Force-exit PIN incorrect.")
            try:
                window.evaluate_js("alert('Incorrect PIN. Exit cancelled.')")
            except:
                pass
            return False

        log("[EXIT] Force-exit PIN accepted. Closing without closing shift.")

        try:
            window.destroy()
        except Exception as e:
            log(f"[EXIT] Error destroying window during force exit: {e}\n{traceback.format_exc()}")

        return True

    # ------------------------------------------------------------
    # 2) Normal exit flow when no shift is active
    # ------------------------------------------------------------
    settings = get_pos_settings()
    require_pin = bool(settings.get("require_pin_exit_program"))

    log(f"[EXIT] require_pin_exit_program={require_pin}")

    # If setting is off, or no PINs exist, allow exit normally.
    if not require_pin or not has_admin_pins():
        log("[EXIT] Exit PIN not required or no admin PINs exist. Closing window.")
        try:
            window.destroy()
        except Exception as e:
            log(f"[EXIT] Error destroying window: {e}\n{traceback.format_exc()}")
        return True

    # Setting is on, so prompt for PIN.
    try:
        pin = window.evaluate_js("prompt('Enter Admin PIN to exit:')")
    except Exception as e:
        log(f"[EXIT] Error prompting for PIN: {e}\n{traceback.format_exc()}")
        return False

    if pin is None:
        log("[EXIT] PIN prompt cancelled.")
        return False

    log("[EXIT] PIN entered; validating...")

    if validate_security_pin(pin):
        log("[EXIT] Correct PIN. Closing window.")
        try:
            window.destroy()
        except Exception as e:
            log(f"[EXIT] Error destroying window: {e}\n{traceback.format_exc()}")
        return True

    log("[EXIT] Incorrect PIN.")
    try:
        window.evaluate_js("alert('Incorrect PIN')")
    except:
        pass

    return False

# -------------------------------------------------------------------
# Auto-update on startup
# -------------------------------------------------------------------
def auto_update_check(window):
    log("[INIT] auto_update_check() starting...")
    url, version = check_for_updates()
    if url:
        log(f"[INIT] Prompting user about update to version {version}")
        window.evaluate_js(f"""
            if (confirm("A new version ({version}) of BrewIns POS is available. Update now?")) {{
                window.pywebview.api.install_update("{url}", "{version}");
            }}
        """)
    else:
        log("[INIT] No update; continuing without prompt.")

# -------------------------------------------------------------------
# MAIN
# -------------------------------------------------------------------
if __name__ == '__main__':
    try:
        log("[MAIN] Launching app_window EXE.")

        # Start Flask
        log("[MAIN] Spawning Flask thread...")
        t = threading.Thread(target=run_flask, daemon=True)
        t.start()

        time.sleep(0.3)
        log("[MAIN] Checking Flask health...")
        if not wait_for_flask():
            log("[MAIN] ERROR: Flask did not start. Showing error window.")
            splash = webview.create_window(
                "Error",
                html="<h1>Flask backend failed to start.</h1><p>See pos_window.log for details.</p>",
                width=500,
                height=300,
                resizable=False,
                easy_drag=False
            )
            webview.start()
            os._exit(1)

        api = ApiBridge()
        log("[MAIN] ApiBridge created.")

        # Splash screen
        log("[MAIN] Creating splash window...")
        splash = webview.create_window(
            "Loading BrewIns POS...",
            url=f"http://{HOST}:{PORT}/splash",
            resizable=False,
            frameless=True,
            fullscreen=True,
            easy_drag=False
        )
        log("[MAIN] Splash window created.")

        def show_main_app():
            log("[MAIN] show_main_app() entered.")
            time.sleep(5)
            log("[MAIN] Destroying splash window...")
            try:
                splash.destroy()
                log("[MAIN] Splash window destroyed.")
            except Exception as e:
                log(f"[MAIN] Error destroying splash: {e}\n{traceback.format_exc()}")
            try:
                webview.settings["ALLOW_DOWNLOADS"] = True
            except:
                pass
                log("[MAIN] Error allowing downloads...")

            log("[MAIN] Creating main POS window...")
            try:
                window = webview.create_window(
                    "BrewIns POS",
                    f"http://{HOST}:{PORT}/pos",
                    resizable=False,
                    frameless=True,
                    fullscreen=True,
                    confirm_close=True,
                    easy_drag=False,
                    js_api=api
                )
                log("[MAIN] Main window created successfully.")
            except Exception as e:
                log(f"[MAIN] ERROR during create_window for POS: {e}\n{traceback.format_exc()}")
                raise

            # On Windows, WebView2 may still allow the frameless window to be
            # moved by dragging the web content. As a robust fallback, start a
            # background thread that finds the native HWND and repeatedly
            # enforces fullscreen position/size so the window cannot be dragged.
            def _get_hwnd(win):
                # Try common attribute names pywebview exposes
                for a in ("hwnd", "_hwnd", "handle", "_handle"):
                    try:
                        h = getattr(win, a)
                        if isinstance(h, int) and h != 0:
                            log(f"[LOCK] Found HWND via attribute {a}: {h}")
                            return h
                    except Exception:
                        pass

                # Fallback: enumerate top-level windows owned by this process.
                try:
                    pid = os.getpid()
                    hwnds = []

                    EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

                    def _enum_proc(hwnd, lParam):
                        pid_out = ctypes.c_ulong()
                        ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid_out))
                        if pid_out.value == pid:
                            # only consider visible top-level windows
                            if ctypes.windll.user32.IsWindowVisible(hwnd):
                                # get title
                                buf = ctypes.create_unicode_buffer(512)
                                ctypes.windll.user32.GetWindowTextW(hwnd, buf, 512)
                                title = buf.value
                                log(f"[LOCK] Enumerated HWND {hwnd} title='{title}' pid={pid_out.value}")
                                hwnds.append((hwnd, title))
                        return True

                    ctypes.windll.user32.EnumWindows(EnumWindowsProc(_enum_proc), 0)
                    if hwnds:
                        # prefer a window whose title matches the expected title
                        title_expected = getattr(win, 'title', None)
                        for h, t in hwnds:
                            if title_expected and title_expected in (t or ""):
                                log(f"[LOCK] Selecting HWND {h} by title match '{t}'")
                                return h
                        # otherwise return the first visible window for this pid
                        h0 = hwnds[0][0]
                        log(f"[LOCK] Selecting first enumerated HWND {h0}")
                        return h0
                except Exception as e:
                    log(f"[LOCK] EnumWindows failed: {e}")

                # Last resort: try to find by window title
                try:
                    title = getattr(win, 'title', None)
                    if title:
                        h = ctypes.windll.user32.FindWindowW(None, title)
                        if h:
                            log(f"[LOCK] Found HWND via FindWindowW: {h}")
                            return h
                except Exception:
                    pass

                return None

            def _lock_fullscreen(win):
                try:
                    if os.name != 'nt':
                        return
                    hwnd = None
                    # Wait a short while for the native handle to appear
                    for _ in range(20):
                        hwnd = _get_hwnd(win)
                        if hwnd:
                            break
                        time.sleep(0.1)
                    if not hwnd:
                        log('[LOCK] Could not obtain HWND for main window.')
                        return

                    SM_CXSCREEN = 0
                    SM_CYSCREEN = 1
                    cx = ctypes.windll.user32.GetSystemMetrics(SM_CXSCREEN)
                    cy = ctypes.windll.user32.GetSystemMetrics(SM_CYSCREEN)

                    GWL_STYLE = -16
                    WS_CAPTION = 0x00C00000
                    WS_THICKFRAME = 0x00040000
                    WS_BORDER = 0x00800000
                    WS_DLGFRAME = 0x00400000
                    WS_SYSMENU = 0x00080000
                    style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_STYLE)
                    new_style = style & ~(WS_CAPTION | WS_THICKFRAME | WS_BORDER | WS_DLGFRAME | WS_SYSMENU)
                    if new_style != style:
                        ctypes.windll.user32.SetWindowLongW(hwnd, GWL_STYLE, new_style)
                        ctypes.windll.user32.SetWindowPos(
                            hwnd,
                            0,
                            0,
                            0,
                            0,
                            0,
                            0x2 | 0x1 | 0x4 | 0x10 | 0x20
                        )
                        log(f'[LOCK] Cleared window move style: {hex(style)} -> {hex(new_style)}')

                    # Install a native hook to prevent client-area dragging and window moving.
                    try:
                        GWL_WNDPROC = -4
                        WM_NCHITTEST = 0x84
                        WM_WINDOWPOSCHANGING = 0x0046
                        WM_MOVING = 0x0216
                        HTCLIENT = 1

                        user32 = ctypes.windll.user32
                        SetWindowLongPtr = getattr(user32, 'SetWindowLongPtrW', user32.SetWindowLongW)

                        ptr_size = ctypes.sizeof(ctypes.c_void_p)
                        LRESULT = ctypes.c_longlong if ptr_size == 8 else ctypes.c_long
                        WPARAM = ctypes.c_void_p
                        LPARAM = ctypes.c_void_p

                        class WINDOWPOS(ctypes.Structure):
                            _fields_ = [
                                ('hwnd', ctypes.c_void_p),
                                ('hwndInsertAfter', ctypes.c_void_p),
                                ('x', ctypes.c_int),
                                ('y', ctypes.c_int),
                                ('cx', ctypes.c_int),
                                ('cy', ctypes.c_int),
                                ('flags', ctypes.c_uint),
                            ]

                        user32.CallWindowProcW.restype = LRESULT
                        user32.CallWindowProcW.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint, WPARAM, LPARAM]
                        call_window_proc = user32.CallWindowProcW

                        old_wndproc = ctypes.c_void_p()

                        def _wnd_proc(hwnd_, msg, wParam, lParam):
                            if msg == WM_NCHITTEST:
                                return HTCLIENT
                            if msg == WM_WINDOWPOSCHANGING and lParam:
                                pos = ctypes.cast(lParam, ctypes.POINTER(WINDOWPOS)).contents
                                if pos.x != 0 or pos.y != 0:
                                    pos.x = 0
                                    pos.y = 0
                                    ctypes.memmove(lParam, ctypes.byref(pos), ctypes.sizeof(pos))
                                return 0
                            if msg == WM_MOVING:
                                return 0
                            return call_window_proc(old_wndproc, hwnd_, msg, wParam, lParam)

                        WNDPROC = ctypes.WINFUNCTYPE(LRESULT, ctypes.c_void_p, ctypes.c_uint, WPARAM, LPARAM)
                        wndproc_ref = WNDPROC(_wnd_proc)
                        SetWindowLongPtr.restype = ctypes.c_void_p
                        SetWindowLongPtr.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_void_p]
                        old_proc = SetWindowLongPtr(hwnd, GWL_WNDPROC, ctypes.cast(wndproc_ref, ctypes.c_void_p))
                        old_wndproc = ctypes.c_void_p(old_proc)
                        log(f'[LOCK] Installed custom WndProc for hwnd={hwnd} old_proc={old_proc}')
                    except Exception as e:
                        log(f'[LOCK] Failed to install WndProc hook: {e}')

                    SWP_NOZORDER = 0x4
                    SWP_NOACTIVATE = 0x10
                    SWP_FRAMECHANGED = 0x20
                    SWP_NOOWNERZORDER = 0x200
                    flags = SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED | SWP_NOOWNERZORDER

                    log(f'[LOCK] Enforcing fullscreen {cx}x{cy} on hwnd={hwnd}')
                    # Keep enforcing until program exit
                    while True:
                        ctypes.windll.user32.SetWindowPos(hwnd, 0, 0, 0, cx, cy, flags)
                        time.sleep(0.1)
                except Exception as e:
                    log(f'[LOCK] Error in lock thread: {e}\n{traceback.format_exc()}')

            try:
                t_lock = threading.Thread(target=_lock_fullscreen, args=(window,), daemon=True)
                t_lock.start()
                log('[LOCK] Started fullscreen lock thread.')
            except Exception as e:
                log(f"[LOCK] Failed to start lock thread: {e}\n{traceback.format_exc()}")

            # ENABLE DOWNLOADS
            try:
                webview.settings["ALLOW_DOWNLOADS"] = True
            except:
                pass
            
            # Attach events
            try:
                window.events.closing += confirm_exit
                log("[MAIN] closing event handler attached.")
            except Exception as e:
                log(f"[MAIN] Failed to attach closing handler: {e}\n{traceback.format_exc()}")

            try:
                def _on_loaded():
                    log("[MAIN] Window loaded event fired; running auto_update_check.")
                    auto_update_check(window)
                window.events.loaded += _on_loaded
                log("[MAIN] loaded event handler attached.")
            except Exception as e:
                log(f"[MAIN] Failed to attach loaded handler: {e}\n{traceback.format_exc()}")

        log("[MAIN] Starting webview event loop...")
        webview.start(show_main_app)
        log("[MAIN] webview.start() returned normally (window closed?).")

    except Exception as e:
        log(f"[FATAL] Unhandled exception in app_window: {e}\n{traceback.format_exc()}")
        # Try to show a basic error window if possible
        try:
            w = webview.create_window(
                "Fatal Error",
                html="<h1>Fatal error in POS app.</h1><p>See pos_window.log for details.</p>",
                width=500,
                height=300
            )
            webview.start()
        except:
            pass
        os._exit(1)
