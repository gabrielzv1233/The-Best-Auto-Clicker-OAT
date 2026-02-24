from logging.handlers import RotatingFileHandler
from rich.logging import RichHandler
import ctypes.wintypes as wintypes
from pathlib import Path
from tkinter import ttk
import tkinter as tk
import threading
import keyboard
import logging
import ctypes
import random
import queue
import time
import copy
import json
import os
import re

DEBUG_MODE = False
TRACE_HOTKEY_EVENTS = False
TRACE_CLICK_EVENTS = False

LOG_TO_FILE = True
LOG_FILE_MAX_BYTES = 1_000_000
LOG_FILE_BACKUPS = 3

UI_THEME_MODE = "system"
WATCH_SYSTEM_THEME = True

APP_NAME = "The Best Auto Clicker OAT"
CONFIG_DIR = Path(os.getenv("APPDATA", str(Path.home()))) / "TheBestAutoClickerOAT"
CONFIG_PATH = CONFIG_DIR / "config.json"
LOG_PATH = CONFIG_DIR / "debug.log"

INPUT_MOUSE = 0
INPUT_KEYBOARD = 1

KEYEVENTF_EXTENDEDKEY = 0x0001
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_SCANCODE = 0x0008

MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
MOUSEEVENTF_RIGHTDOWN = 0x0008
MOUSEEVENTF_RIGHTUP = 0x0010
MOUSEEVENTF_MIDDLEDOWN = 0x0020
MOUSEEVENTF_MIDDLEUP = 0x0040

MAPVK_VSC_TO_VK_EX = 3


def setup_logger():
    logger = logging.getLogger("autoclicker")
    logger.handlers.clear()
    logger.setLevel(logging.DEBUG if DEBUG_MODE else logging.INFO)
    logger.propagate = False

    file_fmt = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(threadName)s | %(message)s",
        datefmt="%H:%M:%S",
    )

    if RichHandler is not None:
        rich_handler = RichHandler(
            rich_tracebacks=True,
            show_time=False,
            show_path=False,
            markup=True,
        )
        rich_handler.setLevel(logging.DEBUG if DEBUG_MODE else logging.INFO)
        rich_handler.setFormatter(logging.Formatter("%(message)s"))
        logger.addHandler(rich_handler)
    else:
        stream_handler = logging.StreamHandler()
        stream_handler.setLevel(logging.DEBUG if DEBUG_MODE else logging.INFO)
        stream_handler.setFormatter(file_fmt)
        logger.addHandler(stream_handler)

    if LOG_TO_FILE:
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            file_handler = RotatingFileHandler(
                LOG_PATH,
                maxBytes=LOG_FILE_MAX_BYTES,
                backupCount=LOG_FILE_BACKUPS,
                encoding="utf-8",
            )
            file_handler.setLevel(logging.DEBUG)
            file_handler.setFormatter(file_fmt)
            logger.addHandler(file_handler)
        except Exception:
            logger.exception("Failed to initialize file logger")

    logger.info("%s logger initialized | DEBUG_MODE=%s", APP_NAME, DEBUG_MODE)
    return logger


logger = setup_logger()

if hasattr(wintypes, "ULONG_PTR"):
    ULONG_PTR = wintypes.ULONG_PTR
else:
    ULONG_PTR = ctypes.c_ulonglong if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_ulong
    logger.debug("wintypes.ULONG_PTR missing, using fallback type: %s", ULONG_PTR)


def default_bind():
    return {
        "name": "",
        "scan_code": None,
        "vk_code": None,
    }


def default_config():
    return {
        "code_display": "hex",
        "cps_mode": "static",
        "static_cps": "12",
        "static_variance": "1",
        "interval_hours": "0",
        "interval_minutes": "0",
        "interval_seconds": "0",
        "interval_milliseconds": "100",
        "output_mode": "mouse",
        "mouse_button": "left",
        "output_key": default_bind(),
        "toggle_mode": "press",
        "start_bind": default_bind(),
        "stop_bind": default_bind(),
    }


def read_windows_app_theme():
    # Windows setting: Settings -> Personalization -> Colors -> "Choose your mode"
    # Registry: AppsUseLightTheme (1=light, 0=dark)
    try:
        import winreg

        key_path = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as k:
            apps_use_light, _ = winreg.QueryValueEx(k, "AppsUseLightTheme")
        return "light" if int(apps_use_light) == 1 else "dark"
    except Exception as e:
        logger.debug("Theme read failed, defaulting to dark | %s", e)
        return "dark"

class WinFocus:
    user32 = ctypes.WinDLL("user32", use_last_error=True)

    class POINT(ctypes.Structure):
        _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]

    class RECT(ctypes.Structure):
        _fields_ = [("left", wintypes.LONG), ("top", wintypes.LONG), ("right", wintypes.LONG), ("bottom", wintypes.LONG)]

    @staticmethod
    def is_cursor_in_window(hwnd: int) -> bool:
        if not hwnd:
            return False

        fg = WinFocus.user32.GetForegroundWindow()
        if fg != hwnd:
            return False

        pt = WinFocus.POINT()
        if not WinFocus.user32.GetCursorPos(ctypes.byref(pt)):
            return False

        rc = WinFocus.RECT()
        if not WinFocus.user32.GetWindowRect(hwnd, ctypes.byref(rc)):
            return False

        return (rc.left <= pt.x <= rc.right) and (rc.top <= pt.y <= rc.bottom)


class WinTimer:
    winmm = ctypes.WinDLL("winmm", use_last_error=True)

    @staticmethod
    def begin(period_ms=1):
        try:
            result = WinTimer.winmm.timeBeginPeriod(int(period_ms))
            if result != 0:
                logger.warning("timeBeginPeriod failed | result=%s", result)
            else:
                logger.info("Timer resolution set to %sms", period_ms)
        except Exception:
            logger.exception("timeBeginPeriod failed")

    @staticmethod
    def     end(period_ms=1):
        try:
            result = WinTimer.winmm.timeEndPeriod(int(period_ms))
            if result != 0:
                logger.warning("timeEndPeriod failed | result=%s", result)
            else:
                logger.info("Timer resolution restored from %sms", period_ms)
        except Exception:
            logger.exception("timeEndPeriod failed")


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", wintypes.LONG),
        ("dy", wintypes.LONG),
        ("mouseData", wintypes.DWORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]


class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [
        ("uMsg", wintypes.DWORD),
        ("wParamL", wintypes.WORD),
        ("wParamH", wintypes.WORD),
    ]


class INPUT_UNION(ctypes.Union):
    _fields_ = [
        ("mi", MOUSEINPUT),
        ("ki", KEYBDINPUT),
        ("hi", HARDWAREINPUT),
    ]


class INPUT(ctypes.Structure):
    _anonymous_ = ("u",)
    _fields_ = [
        ("type", wintypes.DWORD),
        ("u", INPUT_UNION),
    ]


class WinInput:
    user32 = ctypes.WinDLL("user32", use_last_error=True)

    @staticmethod
    def _send(inp):
        sent = WinInput.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))
        if sent != 1:
            last_error = ctypes.get_last_error()
            logger.error("SendInput failed | sent=%s | last_error=%s", sent, last_error)
            return False
        return True

    @staticmethod
    def _send_many(inputs):
        arr_type = INPUT * len(inputs)
        arr = arr_type(*inputs)
        sent = WinInput.user32.SendInput(len(inputs), ctypes.byref(arr), ctypes.sizeof(INPUT))
        if sent != len(inputs):
            last_error = ctypes.get_last_error()
            logger.error("SendInput batch failed | sent=%s/%s | last_error=%s", sent, len(inputs), last_error)
            return False
        return True

    @staticmethod
    def _normalize_scan(scan_code):
        raw = int(scan_code)
        extended = raw > 0xFF
        scan = raw & 0xFF
        return raw, scan, extended

    @staticmethod
    def map_vk(scan_code):
        raw, scan, extended = WinInput._normalize_scan(scan_code)
        scan_for_map = ((0xE0 << 8) | scan) if extended else scan
        vk = WinInput.user32.MapVirtualKeyW(scan_for_map, MAPVK_VSC_TO_VK_EX)
        return int(vk)

    @staticmethod
    def send_key(scan_code, is_keyup=False):
        raw, scan, extended = WinInput._normalize_scan(scan_code)
        flags = KEYEVENTF_SCANCODE
        if extended:
            flags |= KEYEVENTF_EXTENDEDKEY
        if is_keyup:
            flags |= KEYEVENTF_KEYUP

        inp = INPUT()
        inp.type = INPUT_KEYBOARD
        inp.ki = KEYBDINPUT(0, scan, flags, 0, 0)

        if TRACE_CLICK_EVENTS and DEBUG_MODE:
            logger.debug("Send key | scan=0x%X | keyup=%s | extended=%s", raw, is_keyup, extended)

        return WinInput._send(inp)

    @staticmethod
    def tap_key(scan_code):
        raw, scan, extended = WinInput._normalize_scan(scan_code)
        flags_down = KEYEVENTF_SCANCODE | (KEYEVENTF_EXTENDEDKEY if extended else 0)
        flags_up = flags_down | KEYEVENTF_KEYUP

        inp_down = INPUT()
        inp_down.type = INPUT_KEYBOARD
        inp_down.ki = KEYBDINPUT(0, scan, flags_down, 0, 0)

        inp_up = INPUT()
        inp_up.type = INPUT_KEYBOARD
        inp_up.type = INPUT_KEYBOARD
        inp_up.ki = KEYBDINPUT(0, scan, flags_up, 0, 0)

        return WinInput._send_many([inp_down, inp_up])

    @staticmethod
    def click_mouse(button_name):
        if TRACE_CLICK_EVENTS and DEBUG_MODE:
            logger.debug("Click mouse | button=%s", button_name)

        if button_name == "left":
            down_flag, up_flag = MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP
        elif button_name == "right":
            down_flag, up_flag = MOUSEEVENTF_RIGHTDOWN, MOUSEEVENTF_RIGHTUP
        elif button_name == "middle":
            down_flag, up_flag = MOUSEEVENTF_MIDDLEDOWN, MOUSEEVENTF_MIDDLEUP
        else:
            logger.error("Invalid mouse button: %s", button_name)
            return False

        inp_down = INPUT()
        inp_down.type = INPUT_MOUSE
        inp_down.mi = MOUSEINPUT(0, 0, 0, down_flag, 0, 0)

        inp_up = INPUT()
        inp_up.type = INPUT_MOUSE
        inp_up.mi = MOUSEINPUT(0, 0, 0, up_flag, 0, 0)

        return WinInput._send_many([inp_down, inp_up])


class ClickerWorker:
    def __init__(self, config_getter, ui_queue_ref, block_click_check=None):
        self.config_getter = config_getter
        self.ui_queue = ui_queue_ref
        self.block_click_check = block_click_check

        self.shutdown_event = threading.Event()
        self.active_event = threading.Event()
        self.nudge_event = threading.Event()

        self.runtime_dirty = threading.Event()
        self.runtime_dirty.set()
        self.runtime_cache = None

        self._current_cps = None
        self._next_cps_update_t = 0.0

        self._blocked_last = None
        self._variance_tick_s = 0.25

        self.thread = threading.Thread(target=self._loop, name="clicker-worker", daemon=True)
        self.thread.start()
        logger.debug("ClickerWorker started")

    def close(self):
        logger.debug("ClickerWorker closing")
        self.shutdown_event.set()
        self.active_event.clear()
        self.nudge_event.set()
        self.thread.join(timeout=2)
        logger.debug("ClickerWorker closed")

    def nudge(self):
        self.runtime_dirty.set()
        self.nudge_event.set()

    def set_active(self, active, reason="manual"):
        if active:
            self.active_event.set()
            self.ui_queue.put(("status", f"Running ({reason})", "running"))
            logger.debug("Clicker started | [yellow]reason[/yellow]=[magenta]%s[/magenta]", reason)
            self._current_cps = None
            self._next_cps_update_t = 0.0
            self._blocked_last = None
        else:
            self.active_event.clear()
            self.ui_queue.put(("status", "Stopped", "stopped"))
            logger.debug("Clicker stopped | [yellow]reason[/yellow]=[magenta]%s[/magenta]", reason)
            self._current_cps = None
            self._next_cps_update_t = 0.0
            self._blocked_last = None
        self.nudge_event.set()

    def toggle_active(self):
        if self.active_event.is_set():
            self.set_active(False, "toggle bind")
        else:
            self.set_active(True, "toggle bind")

    def _sleep_interruptible(self, seconds_value):
        delay = max(0.0, float(seconds_value))
        if delay <= 0:
            return

        target = time.perf_counter() + delay

        while not self.shutdown_event.is_set():
            if not self.active_event.is_set():
                return

            remaining = target - time.perf_counter()
            if remaining <= 0:
                return

            if self.nudge_event.is_set():
                self.nudge_event.clear()

            if remaining > 0.003:
                time.sleep(remaining - 0.0015)
                continue

            while not self.shutdown_event.is_set():
                if not self.active_event.is_set():
                    return
                if self.nudge_event.is_set():
                    self.nudge_event.clear()
                if time.perf_counter() >= target:
                    return
                time.sleep(0)

    def _loop(self):
        logger.debug("Worker loop entered")
        next_t = None

        while not self.shutdown_event.is_set():
            if not self.active_event.wait(timeout=0.1):
                next_t = None
                self._blocked_last = None
                continue

            if self.runtime_cache is None or self.runtime_dirty.is_set():
                self.runtime_dirty.clear()
                runtime = self.config_getter()
                if not runtime["ok"]:
                    logger.warning("Runtime config invalid: %s", runtime["error"])
                    self.ui_queue.put(("status", f"Config error: {runtime['error']}", "error"))
                    self.active_event.clear()
                    self.runtime_cache = None
                    next_t = None
                    self._current_cps = None
                    self._next_cps_update_t = 0.0
                    self._blocked_last = None
                    continue
                self.runtime_cache = runtime
                next_t = None
                self._current_cps = None
                self._next_cps_update_t = 0.0
                self._blocked_last = None
            else:
                runtime = self.runtime_cache

            blocked = False
            if self.block_click_check is not None and self.block_click_check():
                blocked = True

            if blocked:
                if self._blocked_last is not True:
                    self.ui_queue.put(("status", "Running (blocked: cursor in app)", "running"))
                    self._blocked_last = True
                self._sleep_interruptible(0.02)
                next_t = None
                self._current_cps = None
                self._next_cps_update_t = 0.0
                continue
            else:
                if self._blocked_last is True:
                    self.ui_queue.put(("status", "Running", "running"))
                self._blocked_last = False

            now = time.perf_counter()
            if next_t is None:
                next_t = now

            if runtime["cps_mode"] == "static":
                base = float(runtime["static_cps"])
                var = float(runtime["static_variance"])

                if self._current_cps is None:
                    self._current_cps = base
                    self._next_cps_update_t = now + self._variance_tick_s

                if now >= self._next_cps_update_t:
                    if var > 0:
                        if float(var).is_integer():
                            delta = random.randint(-int(var), int(var))
                            self._current_cps = max(0.001, base + delta)
                        else:
                            delta = random.uniform(-var, var)
                            self._current_cps = max(0.001, base + delta)
                    else:
                        self._current_cps = base

                    self._next_cps_update_t = now + self._variance_tick_s

                    logger.debug("CPS target now: %.3f (base=%.3f var=%.3f)", float(self._current_cps), base, var)

                period = 1.0 / max(0.001, float(self._current_cps))
            else:
                self._current_cps = None
                self._next_cps_update_t = 0.0
                period = float(runtime["interval_seconds"])

            if now < next_t:
                self._sleep_interruptible(next_t - now)
                continue

            if runtime["output_mode"] == "mouse":
                ok = WinInput.click_mouse(runtime["mouse_button"])
            else:
                ok = WinInput.tap_key(runtime["output_key"]["scan_code"])

            if not ok:
                logger.error("Input send failed, stopping clicker")
                self.ui_queue.put(("status", "Input send failed", "error"))
                self.active_event.clear()
                next_t = None
                self._current_cps = None
                self._next_cps_update_t = 0.0
                self._blocked_last = None
                continue

            next_t += period

            after = time.perf_counter()
            if after > next_t + (period * 4):
                next_t = after

        logger.debug("Worker loop exited")


class AutoClickerApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME)
        self.hwnd = int(self.root.winfo_id())

        self.ui_queue = queue.Queue()
        self.config_lock = threading.Lock()
        self.hotkey_lock = threading.Lock()
        self.capture_lock = threading.Lock()

        self.config = self._load_config()
        self.pressed_scans = set()
        self.capture_target = None
        self.capture_threads = {}
        self.current_theme_mode = None
        self.text_input_focused = threading.Event()

        self.style = ttk.Style(self.root)
        self._apply_theme(force=True)

        self.worker = ClickerWorker(
            self._build_runtime_config,
            self.ui_queue,
            block_click_check=lambda: WinFocus.is_cursor_in_window(self.hwnd),
        )

        self.kb_hook = None
        try:
            self.kb_hook = keyboard.hook(self._on_keyboard_event, suppress=False)
            logger.info("Global keyboard hook registered")
        except Exception:
            logger.exception("Failed to register keyboard hook")
            raise

        self.status_var = tk.StringVar(value="Stopped")
        self.status_kind = "stopped"

        self.validation_error = ""
        self.footer_status_var = tk.StringVar(value="Stopped")

        self._trace_guard = False

        self.vars = {}
        self._build_ui()
        self._load_vars_from_config()
        self._apply_state()
        self._refresh_validation()
        self._apply_theme(force=True)

        self.root.update_idletasks()
        w = self.root.winfo_reqwidth()
        h = self.root.winfo_reqheight() + 1
        self.root.geometry(f"{w}x{h}")
        self.root.minsize(w, h)
        self.root.resizable(False, False)

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(50, self._drain_ui_queue)
        if WATCH_SYSTEM_THEME and UI_THEME_MODE == "system":
            self.root.after(1000, self._theme_watch_tick)

        logger.debug("UI initialized")

    def _load_config(self):
        cfg = default_config()
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            if CONFIG_PATH.exists():
                loaded = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
                if isinstance(loaded, dict):
                    self._merge_config(cfg, loaded)
                logger.info("Config loaded from %s", CONFIG_PATH)
            else:
                logger.info("Config file not found, using defaults")
        except Exception:
            logger.exception("Failed to load config, using defaults")
        return cfg

    def _merge_config(self, base, loaded):
        for key, value in loaded.items():
            if key not in base:
                continue
            if isinstance(base[key], dict) and isinstance(value, dict):
                for inner_key, inner_val in value.items():
                    if inner_key in base[key]:
                        base[key][inner_key] = inner_val
            else:
                base[key] = value

    def _save_config(self):
        with self.config_lock:
            data = copy.deepcopy(self.config)
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            temp_path = CONFIG_PATH.with_suffix(".tmp")
            temp_path.write_text(json.dumps(data, indent=2), encoding="utf-8")
            temp_path.replace(CONFIG_PATH)
            logger.debug("Config saved")
        except Exception:
            logger.exception("Failed to save config")

    def _theme_watch_tick(self):
        try:
            self._apply_theme(force=False)
        except Exception:
            logger.exception("Theme watch tick failed")
        self.root.after(1000, self._theme_watch_tick)

    def _apply_theme(self, force=False):
        if UI_THEME_MODE == "system":
            theme_mode = read_windows_app_theme()
        elif UI_THEME_MODE in ("dark", "light"):
            theme_mode = UI_THEME_MODE
        else:
            theme_mode = "dark"
        if not force and theme_mode == self.current_theme_mode:
            return
        self.current_theme_mode = theme_mode
        logger.info("Applying UI theme: %s", theme_mode)

        try:
            self.style.theme_use("clam")
        except Exception:
            logger.exception("Failed to set ttk theme to clam")

        if theme_mode == "dark":
            c = {
                "bg": "#1e1e1e",
                "panel": "#252526",
                "panel2": "#2b2b2b",
                "section_bg": "#1e1e1e",
                "section_border": "#3b3b3b",
                "text": "#f2f2f2",
                "muted": "#c8c8c8",
                "border": "#3b3b3b",
                "field": "#2f2f2f",
                "field_disabled": "#262626",
                "accent": "#4cc2ff",
                "btn": "#313131",
                "btn_hover": "#3a3a3a",
                "ok": "#6ad38b",
                "warn": "#f0c674",
                "err": "#ff6b6b",
            }
        else:
            c = {
                "bg": "#f3f3f3",
                "panel": "#ffffff",
                "panel2": "#fbfbfb",
                "section_bg": "#f3f3f3",
                "section_border": "#d4d4d4",
                "text": "#111111",
                "muted": "#4a4a4a",
                "border": "#d4d4d4",
                "field": "#ffffff",
                "field_disabled": "#efefef",
                "accent": "#0067c0",
                "btn": "#ffffff",
                "btn_hover": "#f5f5f5",
                "ok": "#1f7a3f",
                "warn": "#8a6400",
                "err": "#b00020",
            }

        self.colors = c
        self.root.configure(bg=c["bg"])

        self.style.configure(".", background=c["bg"], foreground=c["text"])
        self.style.configure("TFrame", background=c["bg"])
        self.style.configure("Card.TFrame", background=c["panel"])
        self.style.configure(
            "Section.TFrame",
            background=c["section_bg"],
            bordercolor=c["section_border"],
            relief="solid",
            borderwidth=1,
        )

        self.style.configure(
            "TLabelframe",
            background=c["bg"],
            bordercolor=c["border"],
            relief="solid",
            borderwidth=1,
        )
        self.style.configure("TLabelframe.Label", background=c["bg"], foreground=c["muted"])

        self.style.configure("TLabel", background=c["bg"], foreground=c["text"])
        self.style.configure("Muted.TLabel", background=c["bg"], foreground=c["muted"])

        self.style.configure(
            "TButton",
            background=c["btn"],
            foreground=c["text"],
            bordercolor=c["border"],
            focusthickness=0,
            padding=(10, 6),
        )
        self.style.map(
            "TButton",
            background=[("active", c["btn_hover"]), ("pressed", c["panel2"])],
            foreground=[("disabled", c["muted"])],
        )

        self.style.configure(
            "TEntry",
            fieldbackground=c["field"],
            foreground=c["text"],
            insertcolor=c["text"],
            bordercolor=c["border"],
            lightcolor=c["border"],
            darkcolor=c["border"],
        )
        self.style.map(
            "TEntry",
            fieldbackground=[("disabled", c["field_disabled"])],
            foreground=[("disabled", c["muted"])],
        )

        self.style.configure(
            "TCombobox",
            fieldbackground=c["field"],
            background=c["btn"],
            foreground=c["text"],
            bordercolor=c["border"],
            arrowcolor=c["text"],
        )
        self.style.map(
            "TCombobox",
            fieldbackground=[("readonly", c["field"]), ("disabled", c["field_disabled"])],
            foreground=[("readonly", c["text"]), ("disabled", c["muted"])],
            selectforeground=[("readonly", c["text"])],
            selectbackground=[("readonly", c["field"])],
        )

        self.style.configure("TRadiobutton", background=c["bg"], foreground=c["text"])
        try:
            self.style.map(
                "TRadiobutton",
                background=[("active", c["bg"])],
                foreground=[("active", c["text"]), ("disabled", c["muted"])],
                indicatorcolor=[
                    ("selected", c["accent"]),
                    ("active", c["btn_hover"]),
                    ("disabled", c["field_disabled"]),
                    ("!selected", c["field"]),
                ],
            )
        except tk.TclError:
            self.style.map(
                "TRadiobutton",
                background=[("active", c["bg"])],
                foreground=[("active", c["text"]), ("disabled", c["muted"])],
            )

        self.style.configure("StatusValue.TLabel", background=c["bg"], foreground=c["muted"])
        self.style.configure("StatusRunning.TLabel", background=c["bg"], foreground=c["ok"])
        self.style.configure("StatusStopped.TLabel", background=c["bg"], foreground=c["muted"])
        self.style.configure("StatusError.TLabel", background=c["bg"], foreground=c["err"])

    def _build_ui(self):
        self.root.configure(padx=12, pady=12)

        main = ttk.Frame(self.root)
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1)

        content_col = ttk.Frame(main)
        content_col.grid(row=0, column=0, sticky="nsew")
        content_col.columnconfigure(0, weight=1)

        section_pad = (2, 2)

        cps_frame = ttk.Frame(content_col, style="Section.TFrame")
        cps_frame.grid(row=0, column=0, sticky="ew", pady=section_pad)
        cps_frame.columnconfigure(0, weight=1)
        cps_frame.columnconfigure(1, weight=1)

        self.vars["cps_mode"] = tk.StringVar()
        ttk.Radiobutton(
            cps_frame,
            text="Static CPS",
            value="static",
            variable=self.vars["cps_mode"],
            command=self._on_any_ui_change,
        ).grid(row=0, column=0, sticky="w", padx=10, pady=(8, 6))
        ttk.Radiobutton(
            cps_frame,
            text="Time Definition",
            value="interval",
            variable=self.vars["cps_mode"],
            command=self._on_any_ui_change,
        ).grid(row=0, column=1, sticky="w", padx=10, pady=(8, 6))

        self.static_row = ttk.Frame(cps_frame)
        self.static_row.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 8))
        self.static_row.columnconfigure(1, weight=1)
        self.static_row.columnconfigure(3, weight=1)

        ttk.Label(self.static_row, text="CPS").grid(row=0, column=0, sticky="w")
        self.vars["static_cps"] = tk.StringVar()
        self.static_cps_entry = ttk.Entry(self.static_row, textvariable=self.vars["static_cps"], width=10)
        self.static_cps_entry.grid(row=0, column=1, sticky="ew", padx=(6, 12))

        ttk.Label(self.static_row, text="+/- Variance").grid(row=0, column=2, sticky="w")
        self.vars["static_variance"] = tk.StringVar()
        self.static_var_entry = ttk.Entry(self.static_row, textvariable=self.vars["static_variance"], width=10)
        self.static_var_entry.grid(row=0, column=3, sticky="ew", padx=(6, 0))

        self.interval_row = ttk.Frame(cps_frame)
        self.interval_row.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 8))
        for col in range(8):
            self.interval_row.columnconfigure(col, weight=1 if col % 2 == 1 else 0)

        ttk.Label(self.interval_row, text="H").grid(row=0, column=0, sticky="w")
        self.vars["interval_hours"] = tk.StringVar()
        self.interval_hours_entry = ttk.Entry(self.interval_row, textvariable=self.vars["interval_hours"], width=6)
        self.interval_hours_entry.grid(row=0, column=1, sticky="ew", padx=(4, 8))

        ttk.Label(self.interval_row, text="M").grid(row=0, column=2, sticky="w")
        self.vars["interval_minutes"] = tk.StringVar()
        self.interval_minutes_entry = ttk.Entry(self.interval_row, textvariable=self.vars["interval_minutes"], width=6)
        self.interval_minutes_entry.grid(row=0, column=3, sticky="ew", padx=(4, 8))

        ttk.Label(self.interval_row, text="S").grid(row=0, column=4, sticky="w")
        self.vars["interval_seconds"] = tk.StringVar()
        self.interval_seconds_entry = ttk.Entry(self.interval_row, textvariable=self.vars["interval_seconds"], width=6)
        self.interval_seconds_entry.grid(row=0, column=5, sticky="ew", padx=(4, 8))

        ttk.Label(self.interval_row, text="MS").grid(row=0, column=6, sticky="w")
        self.vars["interval_milliseconds"] = tk.StringVar()
        self.interval_ms_entry = ttk.Entry(self.interval_row, textvariable=self.vars["interval_milliseconds"], width=6)
        self.interval_ms_entry.grid(row=0, column=7, sticky="ew", padx=(4, 0))

        self.text_input_widgets = [
            self.static_cps_entry,
            self.static_var_entry,
            self.interval_hours_entry,
            self.interval_minutes_entry,
            self.interval_seconds_entry,
            self.interval_ms_entry,
        ]
        for widget in self.text_input_widgets:
            widget.bind("<FocusIn>", self._on_text_input_focus_in, add="+")
            widget.bind("<FocusOut>", self._on_text_input_focus_out, add="+")

        output_frame = ttk.Frame(content_col, style="Section.TFrame")
        output_frame.grid(row=1, column=0, sticky="ew", pady=section_pad)
        output_frame.columnconfigure(0, weight=1)
        output_frame.columnconfigure(1, weight=1)

        self.vars["output_mode"] = tk.StringVar()
        ttk.Radiobutton(
            output_frame,
            text="Mouse",
            value="mouse",
            variable=self.vars["output_mode"],
            command=self._on_any_ui_change,
        ).grid(row=0, column=0, sticky="w", padx=10, pady=(8, 6))
        ttk.Radiobutton(
            output_frame,
            text="Keyboard",
            value="keyboard",
            variable=self.vars["output_mode"],
            command=self._on_any_ui_change,
        ).grid(row=0, column=1, sticky="w", padx=10, pady=(8, 6))

        self.mouse_row = ttk.Frame(output_frame)
        self.mouse_row.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 8))
        self.mouse_row.columnconfigure(1, weight=1)

        ttk.Label(self.mouse_row, text="Mouse Button").grid(row=0, column=0, sticky="w")
        self.vars["mouse_button"] = tk.StringVar()
        self.mouse_button_combo = ttk.Combobox(
            self.mouse_row,
            state="readonly",
            textvariable=self.vars["mouse_button"],
            values=["left", "middle", "right"],
            width=14,
        )
        self.mouse_button_combo.grid(row=0, column=1, sticky="w", padx=(8, 0))
        self.mouse_button_combo.bind("<<ComboboxSelected>>", lambda e: self._on_any_ui_change())

        self.keyboard_row = ttk.Frame(output_frame)
        self.keyboard_row.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 8))
        self.keyboard_row.columnconfigure(0, weight=1)

        self.capture_output_button = ttk.Button(
            self.keyboard_row,
            text="Output [Not Set]",
            command=lambda: self._start_capture("output_key"),
        )
        self.capture_output_button.grid(row=0, column=0, sticky="ew")

        binds_frame = ttk.Frame(content_col, style="Section.TFrame")
        binds_frame.grid(row=2, column=0, sticky="ew", pady=section_pad)
        binds_frame.columnconfigure(0, weight=1)
        binds_frame.columnconfigure(1, weight=1)

        toggle_row = ttk.Frame(binds_frame)
        toggle_row.grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(8, 6))

        self.vars["toggle_mode"] = tk.StringVar()
        ttk.Radiobutton(toggle_row, text="Press", value="press", variable=self.vars["toggle_mode"], command=self._on_any_ui_change).grid(row=0, column=0, sticky="w", padx=(0, 12))
        ttk.Radiobutton(toggle_row, text="Toggle", value="toggle", variable=self.vars["toggle_mode"], command=self._on_any_ui_change).grid(row=0, column=1, sticky="w", padx=(0, 12))
        ttk.Radiobutton(toggle_row, text="Separate Stop", value="separate_toggle", variable=self.vars["toggle_mode"], command=self._on_any_ui_change).grid(row=0, column=2, sticky="w")

        self.start_bind_button = ttk.Button(
            binds_frame,
            text="Start [Not Set]",
            command=lambda: self._start_capture("start_bind"),
        )
        self.start_bind_button.grid(row=1, column=0, sticky="ew", padx=(10, 2), pady=(0, 8))

        self.stop_bind_button = ttk.Button(
            binds_frame,
            text="Stop [Not Set]",
            command=lambda: self._start_capture("stop_bind"),
        )
        self.stop_bind_button.grid(row=1, column=1, sticky="ew", padx=(2, 10), pady=(0, 8))

        controls_frame = ttk.Frame(content_col, style="Section.TFrame")
        controls_frame.grid(row=3, column=0, sticky="ew", pady=(16, 2))
        controls_frame.columnconfigure(0, weight=1)
        controls_frame.columnconfigure(1, weight=1)

        self.manual_start_btn = ttk.Button(controls_frame, text="Manual Start", command=lambda: self.worker.set_active(True, "manual"))
        self.manual_start_btn.grid(row=0, column=0, sticky="ew", padx=(10, 2), pady=8)

        self.manual_stop_btn = ttk.Button(controls_frame, text="Manual Stop", command=lambda: self.worker.set_active(False, "manual"))
        self.manual_stop_btn.grid(row=0, column=1, sticky="ew", padx=(2, 10), pady=8)

        footer = ttk.Frame(content_col)
        footer.grid(row=4, column=0, sticky="ew", pady=(2, 0))
        footer.columnconfigure(0, weight=1)

        self.status_label = ttk.Label(footer, textvariable=self.footer_status_var, style="StatusStopped.TLabel")
        self.status_label.grid(row=0, column=0, sticky="w")

        # Right-click to clear binds
        self.start_bind_button.bind("<Button-3>", lambda e: self._clear_bind("start_bind"))
        self.stop_bind_button.bind("<Button-3>", lambda e: self._clear_bind("stop_bind"))
        self.capture_output_button.bind("<Button-3>", lambda e: self._clear_bind("output_key"))

        for key in [
            "static_cps",
            "static_variance",
            "interval_hours",
            "interval_minutes",
            "interval_seconds",
            "interval_milliseconds",
            "cps_mode",
            "output_mode",
            "mouse_button",
            "toggle_mode",
        ]:
            self.vars[key].trace_add("write", self._on_var_trace)

    def _clear_bind(self, target_key):
        with self.capture_lock:
            if self.capture_target is not None:
                return

        with self.config_lock:
            self.config[target_key] = default_bind()

        self._save_config()
        self._refresh_bind_buttons()
        self._refresh_validation()
        self.worker.nudge()
        self._set_status(f"Cleared {target_key.replace('_', ' ')}", "info")
        logger.info("Bind cleared | %s", target_key)

    def _load_vars_from_config(self):
        self._trace_guard = True
        with self.config_lock:
            cfg = copy.deepcopy(self.config)

        for key in [
            "cps_mode",
            "static_cps",
            "static_variance",
            "interval_hours",
            "interval_minutes",
            "interval_seconds",
            "interval_milliseconds",
            "output_mode",
            "mouse_button",
            "toggle_mode",
        ]:
            if key in self.vars:
                self.vars[key].set(str(cfg.get(key, "")))

        self._trace_guard = False
        self._refresh_bind_buttons()

    def _on_var_trace(self, *args):
        self._on_any_ui_change()

    def _on_any_ui_change(self):
        if self._trace_guard:
            return
        self._write_config_from_vars()
        self._save_config()
        self._apply_state()
        self._refresh_validation()
        self.worker.nudge()

    def _write_config_from_vars(self):
        with self.config_lock:
            for key in [
                "cps_mode",
                "static_cps",
                "static_variance",
                "interval_hours",
                "interval_minutes",
                "interval_seconds",
                "interval_milliseconds",
                "output_mode",
                "mouse_button",
                "toggle_mode",
            ]:
                self.config[key] = self.vars[key].get()

    def _set_bind(self, target_key, bind_data):
        with self.config_lock:
            self.config[target_key] = bind_data
        self._save_config()
        self._refresh_bind_buttons()
        self._refresh_validation()
        self.worker.nudge()
        logger.info("Bind updated | %s=%s", target_key, self._format_bind(bind_data))

    def _format_bind(self, bind_data):
        if not bind_data or bind_data.get("scan_code") in (None, ""):
            return "Not Set"
        name = str(bind_data.get("name") or "Key")
        if len(name) == 1:
            display_name = name.upper()
        else:
            display_name = re.sub(r"\s+", " ", name.replace("_", " ")).title()
        scan_code = int(bind_data["scan_code"])
        return f"{display_name} (0x{scan_code:X})"

    def _refresh_bind_buttons(self):
        with self.config_lock:
            cfg = copy.deepcopy(self.config)

        self.start_bind_button.configure(text=f"Start [{self._format_bind(cfg['start_bind'])}]")
        self.stop_bind_button.configure(text=f"Stop [{self._format_bind(cfg['stop_bind'])}]")
        self.capture_output_button.configure(text=f"Output [{self._format_bind(cfg['output_key'])}]")

    def _apply_state(self):
        static_enabled = self.vars["cps_mode"].get() == "static"
        output_mouse = self.vars["output_mode"].get() == "mouse"

        if static_enabled:
            self.static_row.grid()
            self.interval_row.grid_remove()
            self.static_cps_entry.configure(state="normal")
            self.static_var_entry.configure(state="normal")
        else:
            self.static_row.grid_remove()
            self.interval_row.grid()
            self.interval_hours_entry.configure(state="normal")
            self.interval_minutes_entry.configure(state="normal")
            self.interval_seconds_entry.configure(state="normal")
            self.interval_ms_entry.configure(state="normal")

        if output_mouse:
            self.mouse_row.grid()
            self.keyboard_row.grid_remove()
            self.mouse_button_combo.configure(state="readonly")
            self.capture_output_button.configure(state="normal")
        else:
            self.mouse_row.grid_remove()
            self.keyboard_row.grid()
            self.mouse_button_combo.configure(state="disabled")
            self.capture_output_button.configure(state="normal")

    def _parse_float(self, text_value, field_name, minimum=None):
        try:
            value = float(str(text_value).strip())
        except Exception:
            return False, f"{field_name} must be a number"
        if minimum is not None and value < minimum:
            return False, f"{field_name} must be >= {minimum}"
        return True, value

    def _parse_int(self, text_value, field_name, minimum=None):
        try:
            value = int(str(text_value).strip() or "0")
        except Exception:
            return False, f"{field_name} must be an integer"
        if minimum is not None and value < minimum:
            return False, f"{field_name} must be >= {minimum}"
        return True, value

    def _bind_same(self, a, b):
        if not a or not b:
            return False
        return a.get("scan_code") is not None and a.get("scan_code") == b.get("scan_code")

    def _build_runtime_config(self):
        with self.config_lock:
            cfg = copy.deepcopy(self.config)

        start_bind = cfg.get("start_bind") or {}
        stop_bind = cfg.get("stop_bind") or {}
        output_key = cfg.get("output_key") or {}

        if start_bind.get("scan_code") is None:
            return {"ok": False, "error": "Start bind is required"}

        if stop_bind.get("scan_code") is not None and self._bind_same(start_bind, stop_bind):
            return {"ok": False, "error": "Start bind and stop bind cannot be the same"}

        if cfg.get("toggle_mode") == "separate_toggle" and stop_bind.get("scan_code") is None:
            return {"ok": False, "error": "Separate toggle requires a stop bind"}

        if cfg.get("output_mode") == "keyboard":
            if output_key.get("scan_code") is None:
                return {"ok": False, "error": "Output key is required in keyboard mode"}
            if self._bind_same(start_bind, output_key):
                return {"ok": False, "error": "Start bind cannot match output key"}
            if stop_bind.get("scan_code") is not None and self._bind_same(stop_bind, output_key):
                return {"ok": False, "error": "Stop bind cannot match output key"}

        if cfg.get("cps_mode") == "static":
            ok, static_cps = self._parse_float(cfg.get("static_cps", ""), "Static CPS", 0.001)
            if not ok:
                return {"ok": False, "error": static_cps}
            ok, variance = self._parse_float(cfg.get("static_variance", ""), "Variance", 0.0)
            if not ok:
                return {"ok": False, "error": variance}
            return {
                "ok": True,
                "cps_mode": "static",
                "static_cps": static_cps,
                "static_variance": variance,
                "output_mode": cfg.get("output_mode"),
                "mouse_button": cfg.get("mouse_button"),
                "output_key": output_key,
                "toggle_mode": cfg.get("toggle_mode"),
                "start_bind": start_bind,
                "stop_bind": stop_bind,
            }

        ok, h = self._parse_int(cfg.get("interval_hours", ""), "Hours", 0)
        if not ok:
            return {"ok": False, "error": h}
        ok, m = self._parse_int(cfg.get("interval_minutes", ""), "Minutes", 0)
        if not ok:
            return {"ok": False, "error": m}
        ok, s = self._parse_int(cfg.get("interval_seconds", ""), "Seconds", 0)
        if not ok:
            return {"ok": False, "error": s}
        ok, ms = self._parse_int(cfg.get("interval_milliseconds", ""), "Milliseconds", 0)
        if not ok:
            return {"ok": False, "error": ms}

        interval_seconds = (h * 3600) + (m * 60) + s + (ms / 1000.0)
        if interval_seconds <= 0:
            return {"ok": False, "error": "Interval must be greater than 0"}

        return {
            "ok": True,
            "cps_mode": "interval",
            "interval_seconds": interval_seconds,
            "output_mode": cfg.get("output_mode"),
            "mouse_button": cfg.get("mouse_button"),
            "output_key": output_key,
            "toggle_mode": cfg.get("toggle_mode"),
            "start_bind": start_bind,
            "stop_bind": stop_bind,
        }

    def _refresh_validation(self):
        runtime = self._build_runtime_config()
        if runtime["ok"]:
            self.validation_error = ""
        else:
            self.validation_error = runtime["error"]
        self._sync_footer_status()

    def _sync_footer_status(self):
        validation_text = str(self.validation_error or "").strip()
        if validation_text:
            self.footer_status_var.set(validation_text)
            self.status_label.configure(style="StatusError.TLabel")
            return

        self.footer_status_var.set(self.status_var.get())
        if self.status_kind == "running":
            self.status_label.configure(style="StatusRunning.TLabel")
        elif self.status_kind in ("error", "warn"):
            self.status_label.configure(style="StatusError.TLabel")
        elif self.status_kind == "stopped":
            self.status_label.configure(style="StatusStopped.TLabel")
        else:
            self.status_label.configure(style="StatusValue.TLabel")

    def _start_capture(self, target_key):
        with self.capture_lock:
            if self.capture_target is not None:
                logger.debug("Capture already in progress, ignoring new capture request")
                return
            self.capture_target = target_key

        logger.info("Starting key capture for %s", target_key)

        if target_key == "start_bind":
            self.start_bind_button.configure(text="Start [...]")
        elif target_key == "stop_bind":
            self.stop_bind_button.configure(text="Stop [...]")
        else:
            self.capture_output_button.configure(text="Output [...]")

        t = threading.Thread(
            target=self._capture_key_thread,
            args=(target_key,),
            daemon=True,
            name=f"capture-{target_key}",
        )
        self.capture_threads[target_key] = t
        t.start()

    def _capture_key_thread(self, target_key):
        try:
            while True:
                event = keyboard.read_event(suppress=True)
                if getattr(event, "event_type", None) != "down":
                    continue

                scan_code = getattr(event, "scan_code", None)
                if scan_code is None:
                    continue

                name = getattr(event, "name", "") or "key"
                vk_code = WinInput.map_vk(int(scan_code))

                bind_data = {
                    "name": str(name),
                    "scan_code": int(scan_code),
                    "vk_code": int(vk_code),
                }

                logger.debug(
                    "Captured key | target=%s | name=%s | scan=0x%X | vk=0x%X",
                    target_key,
                    bind_data["name"],
                    bind_data["scan_code"],
                    bind_data["vk_code"],
                )
                self.ui_queue.put(("capture_done", target_key, bind_data))
                break
        except Exception:
            logger.exception("Capture thread failed for %s", target_key)
            self.ui_queue.put(("capture_error", target_key, "Capture failed, check logs"))

    def _finish_capture_ui(self, target_key):
        with self.capture_lock:
            self.capture_target = None
        self.capture_threads.pop(target_key, None)
        self._refresh_bind_buttons()

    def _set_status(self, text, kind="info"):
        self.status_var.set(text)
        self.status_kind = kind
        self._sync_footer_status()

    def _drain_ui_queue(self):
        try:
            while True:
                item = self.ui_queue.get_nowait()
                action = item[0]

                if action == "status":
                    text = item[1]
                    kind = item[2] if len(item) > 2 else "info"
                    self._set_status(text, kind)

                elif action == "capture_done":
                    _, target_key, bind_data = item
                    self._finish_capture_ui(target_key)
                    self._set_bind(target_key, bind_data)
                    self._set_status(f"Captured {self._format_bind(bind_data)}", "info")

                elif action == "capture_error":
                    _, target_key, err_text = item
                    self._finish_capture_ui(target_key)
                    self._set_status(f"Capture error: {err_text}", "error")

        except queue.Empty:
            pass
        except Exception:
            logger.exception("UI queue processing error")

        self.root.after(50, self._drain_ui_queue)

    def _on_text_input_focus_in(self, _event=None):
        self.text_input_focused.set()

    def _on_text_input_focus_out(self, _event=None):
        self.root.after_idle(self._refresh_text_input_focus_state)

    def _refresh_text_input_focus_state(self):
        focused_widget = self.root.focus_get()
        if focused_widget is None:
            self.text_input_focused.clear()
            return
        try:
            widget_class = str(focused_widget.winfo_class())
        except Exception:
            self.text_input_focused.clear()
            return
        if widget_class in ("Entry", "TEntry", "Spinbox", "Text"):
            self.text_input_focused.set()
        else:
            self.text_input_focused.clear()

    def _on_keyboard_event(self, event):
        with self.capture_lock:
            if self.capture_target is not None:
                return

        try:
            scan_code = int(getattr(event, "scan_code", -1))
        except Exception:
            return

        event_type = getattr(event, "event_type", "")
        if event_type not in ("down", "up"):
            return

        is_down = event_type == "down"

        with self.config_lock:
            start_scan = self.config["start_bind"].get("scan_code")
            stop_scan = self.config["stop_bind"].get("scan_code")
            toggle_mode = self.config.get("toggle_mode", "press")

        if TRACE_HOTKEY_EVENTS and DEBUG_MODE:
            logger.debug(
                "Hotkey event | type=%s | scan=0x%X | start=%s | stop=%s | mode=%s",
                event_type,
                scan_code,
                None if start_scan is None else f"0x{int(start_scan):X}",
                None if stop_scan is None else f"0x{int(stop_scan):X}",
                toggle_mode,
            )

        with self.hotkey_lock:
            if is_down:
                if scan_code in self.pressed_scans:
                    return
                self.pressed_scans.add(scan_code)
            else:
                if scan_code in self.pressed_scans:
                    self.pressed_scans.remove(scan_code)

        if stop_scan is not None and scan_code == int(stop_scan) and is_down:
            self.worker.set_active(False, "stop bind")
            return

        if start_scan is None or scan_code != int(start_scan):
            return

        ignore_start_for_window = is_down and WinFocus.is_cursor_in_window(self.hwnd)
        ignore_start_for_text_input = is_down and self.text_input_focused.is_set()
        if ignore_start_for_window or ignore_start_for_text_input:
            if ignore_start_for_text_input:
                logger.debug("Start hotkey ignored (text input focused)")
            else:
                logger.debug("Start hotkey ignored (app in foreground and cursor inside window)")
            return

        if toggle_mode == "press":
            if is_down:
                self.worker.set_active(True, "hold bind")
            else:
                self.worker.set_active(False, "hold bind")
            return

        if toggle_mode == "toggle":
            if is_down:
                self.worker.toggle_active()
            return

        if toggle_mode == "separate_toggle":
            if is_down:
                self.worker.set_active(True, "stop bind")
            return

    def _on_close(self):
        logger.info("Closing app")
        try:
            if self.kb_hook is not None:
                keyboard.unhook(self.kb_hook)
                logger.debug("Keyboard hook removed")
        except Exception:
            logger.exception("Failed to unhook keyboard")
        self.worker.close()
        self.root.destroy()


def main():
    logger.info("Launching %s", APP_NAME)
    WinTimer.begin(1)

    root = tk.Tk()
    return root, AutoClickerApp(root)

if __name__ == "__main__":
    root = None
    app = None
    try:
        root, app = main()
        root.mainloop()
    except KeyboardInterrupt:
        logger.info("Interrupted by user, exiting")
        if app is not None:
            app._on_close()
    finally:
        WinTimer.end(1)
