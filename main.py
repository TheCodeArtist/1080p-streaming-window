"""
1080p Window Resizer
A system-tray utility that forces any chosen window to 1920×1080,
ensuring a clean 1080p capture for Twitch / OBS window-capture sources.
"""

import ctypes
import os
import sys
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import ttk, messagebox

import win32gui
import win32con
import win32api
import win32ui
import win32process

# Tell Windows this process handles its own DPI scaling so all Win32
# coordinates are in physical pixels, not logical (scaled) pixels.
# WHY: Without this, Windows would lie to us about window sizes on high-DPI screens (e.g., 150% scale),
# causing our 1920x1080 resize to result in a physically smaller or larger window than intended.
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)   # Per-monitor DPI aware
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()    # System DPI aware (fallback)
    except Exception:
        pass
import pystray
from PIL import Image, ImageTk

# ── App icon ─────────────────────────────────────────────────────────────────
def _resource_path(filename: str) -> Path:
    """Resolve a bundled resource path - works both frozen (PyInstaller) and from source.
    
    WHY: PyInstaller unpacks data to a temporary folder (_MEIPASS) at runtime. 
    We must check for this attribute to find assets when running as a compiled .exe.
    """
    base = Path(sys._MEIPASS) if hasattr(sys, "_MEIPASS") else Path(__file__).parent
    return base / filename

_ICON_PATH = _resource_path("1080p.png")

_app_icon_base: "Image.Image | None" = None


def _load_app_icon(size: int | tuple[int, int] | None = None) -> Image.Image:
    """Load (and cache) the bundled 1080p.png, returning a copy resized to *size*."""
    global _app_icon_base
    if _app_icon_base is None:
        _app_icon_base = Image.open(_ICON_PATH).convert("RGBA")
    img = _app_icon_base.copy()
    if size is not None:
        if isinstance(size, int):
            size = (size, size)
        img = img.resize(size, Image.LANCZOS)
    return img

# ── Constants ────────────────────────────────────────────────────────────────
TARGET_W = 1920
TARGET_H = 1080

# Light-mode palette - all foreground/background pairs meet WCAG AA (4.5:1+)
APP_BG       = "#f0f4f8"   # cool blue-grey app shell
PANEL_BG     = "#ffffff"   # card / panel surfaces
ROW_ALT      = "#f7f9fc"   # alternating list row tint
ACCENT       = "#2563eb"   # blue-600  - primary action  (8.3:1 on white)
ACCENT_HOVER = "#1d4ed8"   # blue-700  - hover / pressed
TEXT_PRIMARY = "#111827"   # grey-900  - headings & body  (18:1 on white)
TEXT_MUTED   = "#6b7280"   # grey-500  - secondary text   (4.6:1 on white)
SUCCESS      = "#15803d"   # green-700 - confirmed / ok   (7.2:1 on white)
WARNING      = "#b45309"   # amber-700 - caution / diff   (4.8:1 on white)
BORDER       = "#e5e7eb"   # grey-200  - subtle dividers
BTN_DISABLED_BG = "#d1d5db"   # grey-300 - disabled button fill
BTN_DISABLED_FG = "#374151"   # grey-700 - 7.8:1 on grey-300, always readable


# ── Tray icon ─────────────────────────────────────────────────────────────────
def _make_tray_image() -> Image.Image:
    return _load_app_icon(64)


# ── Win32 helpers ─────────────────────────────────────────────────────────────

# Shell/desktop window classes that are never useful capture targets
# WHY: These are internal Windows components (Taskbar, Desktop, Start Menu) 
# that users never want to resize for streaming.
_SHELL_CLASSES = frozenset({
    "Progman",          # Windows desktop ("Program Manager")
    "WorkerW",          # Desktop wallpaper worker
    "Shell_TrayWnd",    # Taskbar
    "Shell_SecondaryTrayWnd",  # Secondary-monitor taskbar
    "DV2ControlHost",   # Start menu host
    "Windows.UI.Core.CoreWindow",  # UWP shell chrome
})

_DWMWA_CLOAKED = 14  # DwmGetWindowAttribute attribute for cloaked state


def _is_real_app_window(hwnd: int, own_pid: int) -> bool:
    """Return True only for windows that belong in a task-switcher list."""
    if not win32gui.IsWindowVisible(hwnd):
        return False

    title = win32gui.GetWindowText(hwnd)
    if not title.strip():
        return False

    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    if pid == own_pid:
        return False

    # Skip shell/desktop windows by class name
    if win32gui.GetClassName(hwnd) in _SHELL_CLASSES:
        return False

    # Skip tool windows (floating palettes, notification pop-ups, …) unless
    # the app explicitly opts them back in with WS_EX_APPWINDOW.
    # WHY: Tool windows don't appear in Alt+Tab, so they shouldn't appear in our list.
    exstyle = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
    if (exstyle & win32con.WS_EX_TOOLWINDOW) and not (exstyle & win32con.WS_EX_APPWINDOW):
        return False

    # Skip cloaked windows (UWP apps on other virtual desktops, etc.)
    # WHY: Windows 10+ "cloaks" windows that are suspended or on other virtual desktops.
    # They are technically "visible" (IsWindowVisible is True) but not shown to the user.
    cloaked = ctypes.c_int(0)
    try:
        ctypes.windll.dwmapi.DwmGetWindowAttribute(
            hwnd, _DWMWA_CLOAKED, ctypes.byref(cloaked), ctypes.sizeof(cloaked)
        )
    except Exception:
        pass
    if cloaked.value:
        return False

    return True


def _enum_windows() -> list[tuple[int, str]]:
    """Return visible top-level windows that have a non-empty title, excluding our own process."""
    own_pid = os.getpid()
    results: list[tuple[int, str]] = []

    def _cb(hwnd, _):
        if _is_real_app_window(hwnd, own_pid):
            results.append((hwnd, win32gui.GetWindowText(hwnd)))

    win32gui.EnumWindows(_cb, None)
    return sorted(results, key=lambda x: x[1].lower())


# ── Window icon helper ────────────────────────────────────────────────────────
_GCL_HICON   = -14
_GCL_HICONSM = -34


def _get_window_icon(hwnd: int, size: int = 20) -> Image.Image | None:
    """Return the window's small icon as a PIL Image, or None on failure.
    
    WHY: Windows stores icons in multiple places. We try them in order of preference:
    1. WM_GETICON (explicitly set by app)
    2. GCL_HICONSM (class small icon)
    3. GCL_HICON (class large icon)
    """
    try:
        hicon = win32gui.SendMessage(hwnd, win32con.WM_GETICON, 0, 0)  # ICON_SMALL
        if not hicon:
            hicon = win32gui.SendMessage(hwnd, win32con.WM_GETICON, 1, 0)  # ICON_BIG
        if not hicon:
            hicon = win32gui.GetClassLong(hwnd, _GCL_HICONSM)
        if not hicon:
            hicon = win32gui.GetClassLong(hwnd, _GCL_HICON)
        if not hicon:
            return None

        # Draw the HICON into a memory DC to convert it to a PIL Image.
        # WHY: HICONs are GDI objects; we need raw pixel data for PIL/Tkinter.
        screen_hdc = win32gui.GetDC(0)
        hdc_mem: win32ui.CDC | None = None
        hbmp: win32ui.PyCBitmap | None = None
        try:
            hdc_screen = win32ui.CreateDCFromHandle(screen_hdc)
            hdc_mem    = hdc_screen.CreateCompatibleDC()
            hbmp       = win32ui.CreateBitmap()
            hbmp.CreateCompatibleBitmap(hdc_screen, size, size)
            hdc_mem.SelectObject(hbmp)
            hdc_mem.FillSolidRect((0, 0, size, size), win32api.RGB(245, 245, 245))
            ctypes.windll.user32.DrawIconEx(
                hdc_mem.GetHandleOutput(), 0, 0, hicon, size, size, 0, None, 3
            )
            bmpstr = hbmp.GetBitmapBits(True)
            img = Image.frombuffer("RGBA", (size, size), bmpstr, "raw", "BGRA", 0, 1)
        finally:
            if hdc_mem is not None:
                hdc_mem.DeleteDC()
            if hbmp is not None:
                win32gui.DeleteObject(hbmp.GetHandle())
            win32gui.ReleaseDC(0, screen_hdc)

        return img
    except Exception:
        return None


class _RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

_DWMWA_EXTENDED_FRAME_BOUNDS = 9


def _get_window_rect(hwnd: int) -> tuple[int, int, int, int]:
    """Outer rect (includes DWM drop-shadow). Returns (left, top, width, height).
    
    WHY: GetWindowRect returns the "window bounds" which, since Vista/7, includes
    invisible drop-shadow areas. Moving a window to (0,0) using these coords
    often leaves a visible gap at the screen edge.
    """
    try:
        l, t, r, b = win32gui.GetWindowRect(hwnd)
        return l, t, r - l, b - t
    except Exception:
        return 0, 0, 0, 0


def _get_visible_rect(hwnd: int) -> tuple[int, int, int, int]:
    """Visible (DWM) rect without drop-shadow. Returns (left, top, width, height).
    
    WHY: DwmGetWindowAttribute with EXTENDED_FRAME_BOUNDS gives the actual
    visible pixels of the window. We use this to calculate the true visual position.
    """
    try:
        rc = _RECT()
        ctypes.windll.dwmapi.DwmGetWindowAttribute(
            hwnd, _DWMWA_EXTENDED_FRAME_BOUNDS,
            ctypes.byref(rc), ctypes.sizeof(rc))
        return rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top
    except Exception:
        return _get_window_rect(hwnd)


def _get_shadow_margins(hwnd: int) -> tuple[int, int, int, int]:
    """(left, top, right, bottom) shadow padding included in GetWindowRect."""
    try:
        ol, ot, or_, ob = win32gui.GetWindowRect(hwnd)
        rc = _RECT()
        ctypes.windll.dwmapi.DwmGetWindowAttribute(
            hwnd, _DWMWA_EXTENDED_FRAME_BOUNDS,
            ctypes.byref(rc), ctypes.sizeof(rc))
        return (rc.left - ol, rc.top - ot, or_ - rc.right, ob - rc.bottom)
    except Exception:
        return (0, 0, 0, 0)


def _get_nonclient_margins(hwnd: int) -> tuple[int, int, int, int]:
    """Return (left, top, right, bottom) non-client thickness in physical pixels.

    'top' includes the title bar plus any top frame.
    Computed by diffing the DWM visible rect against the client area in screen
    coordinates, so it automatically captures borders of any thickness.
    Returns (0,0,0,0) for borderless/fullscreen windows.
    
    WHY: We can't just ask for "title bar height" because themes and custom
    window frames (like Chrome/Electron) make standard metrics unreliable.
    Comparing the visible outer rect to the inner client rect is the only robust way.
    """
    try:
        vx, vy, vw, vh = _get_visible_rect(hwnd)
        cl, ct, cr, cb = win32gui.GetClientRect(hwnd)
        # map client corners to screen coordinates
        cx0, cy0 = win32gui.ClientToScreen(hwnd, (cl, ct))
        cx1, cy1 = win32gui.ClientToScreen(hwnd, (cr, cb))
        return (
            max(0, cx0 - vx),           # left border
            max(0, cy0 - vy),           # title bar + top frame
            max(0, (vx + vw) - cx1),    # right border
            max(0, (vy + vh) - cy1),    # bottom border
        )
    except Exception:
        return (0, 0, 0, 0)


def _get_monitor_top_left(hwnd: int) -> tuple[int, int]:
    """Return the top-left (x, y) of the monitor that contains the given window."""
    try:
        monitor = win32api.MonitorFromWindow(hwnd, win32con.MONITOR_DEFAULTTONEAREST)
        info = win32api.GetMonitorInfo(monitor)
        ml, mt, _, _ = info["Monitor"]
        return ml, mt
    except Exception:
        return 0, 0


def _dpi_awareness_str() -> str:
    try:
        v = ctypes.windll.shcore.GetProcessDpiAwareness(0)
        return {0: "UNAWARE", 1: "SYSTEM", 2: "PER_MONITOR"}.get(v, str(v))
    except Exception:
        return "unknown"


def _window_dpi(hwnd: int) -> int:
    try:
        return ctypes.windll.user32.GetDpiForWindow(hwnd)
    except Exception:
        return -1


def _get_client_size(hwnd: int) -> tuple[int, int]:
    """Client (content) area dimensions in physical pixels."""
    ncl, nct, ncr, ncb = _get_nonclient_margins(hwnd)
    _, _, vw, vh = _get_visible_rect(hwnd)
    return vw - ncl - ncr, vh - nct - ncb


def _client_wh(hwnd: int) -> tuple[int, int]:
    """Read client area via GetClientRect (no NC-margin inference)."""
    cl, ct, cr, cb = win32gui.GetClientRect(hwnd)
    return cr - cl, cb - ct


def _resize_window(hwnd: int, move_to_origin: bool,
                   log_fn=None) -> str:
    """Resize (and optionally reposition) the target window so the client
    (content) area is exactly TARGET_W × TARGET_H pixels.

    Strategy: measure outer-rect and client-rect *after* the window has
    settled into its restored state, then compute the required outer size
    as a pure delta (target_client - current_client + current_outer).
    This avoids NC-margin inference errors that arise when the window is
    maximized or in a transitional state.  A single retry after a short
    sleep handles apps whose WM_SIZE handler runs asynchronously on a
    render/game thread.
    """
    if log_fn is None:
        log_fn = lambda _: None  # discard by default
    try:
        title    = win32gui.GetWindowText(hwnd)
        win_dpi  = _window_dpi(hwnd)
        proc_dpi = _dpi_awareness_str()
        style    = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        is_max   = bool(style & win32con.WS_MAXIMIZE)

        # WHY: You cannot reliably resize a maximized window. It must be restored first.
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        if is_max:
            # Give the window (and any async WM_SIZE handler) time to settle
            # into the fully-restored state before we measure geometry.
            time.sleep(0.10)

        # Re-measure everything *after* restore is complete.
        lx, ly, lw, lh = _get_window_rect(hwnd)
        vx, vy, vw, vh = _get_visible_rect(hwnd)
        sl, st, sr, sb = _get_shadow_margins(hwnd)
        ncl, nct, ncr, ncb = _get_nonclient_margins(hwnd)
        cw, ch = _client_wh(hwnd)  # authoritative: straight from GetClientRect

        log_fn(f"── Resize ───────────────────────────────────────")
        log_fn(f"  Window   : {title!r} (hwnd={hwnd:#010x})")
        log_fn(f"  Outer    : pos=({lx},{ly})  size={lw}×{lh}")
        log_fn(f"  Visible  : pos=({vx},{vy})  size={vw}×{vh}")
        log_fn(f"  Shadow   : L={sl} T={st} R={sr} B={sb}")
        log_fn(f"  NC frame : L={ncl} T={nct} R={ncr} B={ncb}  "
               f"(title bar={nct}px, border L={ncl} R={ncr} B={ncb})")
        log_fn(f"  Client   : {cw}×{ch}  (GetClientRect)")
        log_fn(f"  DPI      : window={win_dpi}  process={proc_dpi}")
        log_fn(f"  Style    : {style:#010x}  maximized={is_max}")

        # Delta approach: required outer = current outer + (target - current) client.
        # This is exact regardless of NC / shadow metric precision because the
        # outer↔client delta is measured from the live window, not inferred.
        # WHY: Calculating "border width" and adding it to 1920 is prone to off-by-one errors.
        # Adding the *difference* between current and target preserves existing frame metrics perfectly.
        adj_w = lw + (TARGET_W - cw)
        adj_h = lh + (TARGET_H - ch)

        # SWP_FRAMECHANGED forces Windows to recompute NC metrics after the move.
        flags = win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE | win32con.SWP_FRAMECHANGED
        if move_to_origin:
            mx, my = _get_monitor_top_left(hwnd)
            # Place the *visible* rect at (mx, my); shadow may extend off-screen.
            # WHY: If we just moved to (0,0), the window would look shifted because of the invisible left shadow.
            x, y = mx - sl, my - st
        else:
            x, y = 0, 0
            flags |= win32con.SWP_NOMOVE

        log_fn(f"  Calling  : SetWindowPos(w={adj_w}, h={adj_h})  "
               f"[Δclient w={TARGET_W - cw:+d}  h={TARGET_H - ch:+d}]")
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, adj_w, adj_h, flags)

        # Wait for any async WM_SIZE handlers (game render threads, etc.) to run.
        time.sleep(0.15)
        fw, fh = _client_wh(hwnd)

        if fw != TARGET_W or fh != TARGET_H:
            # One retry: re-measure the settled state and apply a fresh delta.
            # WHY: Some apps (like games or Electron apps) clamp their size or snap to a grid
            # on the first resize event. A second pass often corrects this.
            lw2, lh2 = _get_window_rect(hwnd)[2:4]
            adj_w2 = lw2 + (TARGET_W - fw)
            adj_h2 = lh2 + (TARGET_H - fh)
            log_fn(f"  Retry    : client={fw}×{fh} → SetWindowPos(w={adj_w2}, h={adj_h2})")
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, adj_w2, adj_h2, flags)
            time.sleep(0.15)
            fw, fh = _client_wh(hwnd)

        _, _, fvw, fvh = _get_visible_rect(hwnd)
        log_fn(f"  After    : visible={fvw}×{fvh}  client={fw}×{fh}")
        log_fn(f"────────────────────────────────────────────────")

        if fw == TARGET_W and fh == TARGET_H:
            return f"Client area {TARGET_W}×{TARGET_H} ✓"
        else:
            return f"Attempted - client {fw}×{fh} (window may have constraints)"
    except Exception as exc:
        log_fn(f"  ERROR   : {exc}")
        return f"Error: {exc}"


# ── Selector window ───────────────────────────────────────────────────────────
class SelectorWindow:
    def __init__(self, root: tk.Tk, on_close=None):
        self._root = root
        self._hwnd_map: dict[str, int] = {}
        self._win: tk.Toplevel | None = None
        self._icon_cache: list[ImageTk.PhotoImage] = []  # prevent GC of PhotoImages
        self._on_close = on_close

    # called from pystray thread → schedule on tk main thread
    # WHY: Tkinter is not thread-safe. All GUI operations must happen on the main thread.
    # pystray runs on a background thread, so we use root.after(0, ...) to dispatch work to the main loop.
    def show(self):
        self._root.after(0, self._build_or_raise)

    def _build_or_raise(self):
        if self._win and self._win.winfo_exists():
            self._win.deiconify()   # un-hide if previously withdrawn
            self._win.lift()
            self._win.focus_force()
            return
        self._build()

    def _build(self):
        win = tk.Toplevel(self._root)
        win.withdraw()  # keep hidden until centred to avoid position flash
        # WHY: Constructing the window while hidden prevents the user from seeing it "jump"
        # from the top-left corner to the center of the screen.
        win.title("1080p Window Resizer")
        win.configure(bg=APP_BG)
        win.resizable(False, False)
        win.protocol("WM_DELETE_WINDOW", self._on_close_to_tray)
        self._win = win

        self._apply_style()

        # ── Header ───────────────────────────────────────────────────────────
        hdr = tk.Frame(win, bg=PANEL_BG, pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="  1080p Window Resizer",
                 font=("Segoe UI", 14, "bold"),
                 fg=ACCENT, bg=PANEL_BG).pack(side="left", padx=8)
        tk.Label(hdr, text="Force any window to 1920×1080",
                 font=("Segoe UI", 9), fg=TEXT_MUTED, bg=PANEL_BG).pack(side="left")

        sep = tk.Frame(win, bg=BORDER, height=1)
        sep.pack(fill="x")

        # ── Window list ──────────────────────────────────────────────────────
        list_frame = tk.Frame(win, bg=APP_BG, padx=14, pady=10)
        list_frame.pack(fill="both", expand=True)

        tk.Label(list_frame, text="Select one or more windows:",
                 font=("Segoe UI", 10, "bold"),
                 fg=TEXT_PRIMARY, bg=APP_BG).pack(anchor="w")
        tk.Label(list_frame, text="Ctrl+click / Shift+click for multi-select · Double-click to resize instantly",
                 font=("Segoe UI", 9), fg=TEXT_MUTED, bg=APP_BG).pack(anchor="w", pady=(0, 6))

        lb_frame = tk.Frame(list_frame, bg=BORDER, padx=1, pady=1)
        lb_frame.pack(fill="both", expand=True)

        scrollbar = ttk.Scrollbar(lb_frame, orient="vertical",
                                  style="Vertical.TScrollbar")
        self._tree = ttk.Treeview(
            lb_frame,
            style="Window.Treeview",
            show="tree",
            selectmode="extended",
            height=9,
            yscrollcommand=scrollbar.set,
        )
        self._tree.column("#0", width=560, minwidth=300, stretch=True)
        # WHY: We use a Treeview with only column "#0" (the tree column) to display
        # the icon and text together. This mimics a rich listbox.
        scrollbar.config(command=self._tree.yview)
        scrollbar.pack(side="right", fill="y")
        self._tree.pack(side="left", fill="both", expand=True)
        self._tree.bind("<Double-1>", lambda _: self._do_resize())
        self._tree.bind("<<TreeviewSelect>>", self._on_select)

        # ── Info bar ─────────────────────────────────────────────────────────
        info_frame = tk.Frame(win, bg=PANEL_BG, padx=14, pady=6)
        info_frame.pack(fill="x")
        tk.Label(info_frame, text="Current size:",
                 font=("Segoe UI", 9), fg=TEXT_MUTED, bg=PANEL_BG).pack(side="left")
        self._size_var = tk.StringVar(value="- select a window -")
        tk.Label(info_frame, textvariable=self._size_var,
                 font=("Segoe UI", 9, "bold"), fg=TEXT_PRIMARY, bg=PANEL_BG).pack(side="left", padx=6)
        self._count_var = tk.StringVar(value="")
        tk.Label(info_frame, textvariable=self._count_var,
                 font=("Segoe UI", 8), fg=TEXT_MUTED, bg=PANEL_BG).pack(side="right")

        sep2 = tk.Frame(win, bg=BORDER, height=1)
        sep2.pack(fill="x")

        # ── Options ──────────────────────────────────────────────────────────
        opt_frame = tk.Frame(win, bg=APP_BG, padx=14, pady=8)
        opt_frame.pack(fill="x")
        self._move_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_frame, text="Also move window to top-left of active display",
                        variable=self._move_var,
                        style="Switch.TCheckbutton").pack(side="left")

        # ── Buttons ──────────────────────────────────────────────────────────
        btn_frame = tk.Frame(win, bg=APP_BG)
        btn_frame.pack(fill="x", padx=14, pady=(0, 14))

        self._status_var = tk.StringVar(value="")
        self._status_lbl = tk.Label(btn_frame, textvariable=self._status_var,
                                    font=("Segoe UI", 9, "italic"),
                                    fg=SUCCESS, bg=APP_BG)
        self._status_lbl.pack(side="left", pady=4)

        self._resize_btn = tk.Button(
            btn_frame,
            text="Resize to 1920 × 1080",
            font=("Segoe UI", 10, "bold"),
            bg=BTN_DISABLED_BG, fg=BTN_DISABLED_FG,
            activebackground=ACCENT_HOVER, activeforeground="#ffffff",
            relief="flat", padx=16, pady=8, cursor="",
            command=self._do_resize,
            state="disabled",
        )
        self._resize_btn.pack(side="right")

        refresh_btn = tk.Button(
            btn_frame,
            text="⟳  Refresh list",
            font=("Segoe UI", 9),
            bg=PANEL_BG, fg=TEXT_PRIMARY,
            activebackground=BORDER, activeforeground=TEXT_PRIMARY,
            relief="flat", padx=10, pady=8, cursor="hand2",
            command=self._populate,
        )
        refresh_btn.pack(side="right", padx=(0, 8))

        # ── Log section ──────────────────────────────────────────────────────
        sep3 = tk.Frame(win, bg=BORDER, height=1)
        sep3.pack(fill="x")

        self._log_expanded = False
        log_outer = tk.Frame(win, bg=APP_BG, padx=14, pady=4)
        log_outer.pack(fill="x")

        log_hdr = tk.Frame(log_outer, bg=APP_BG)
        log_hdr.pack(fill="x")

        self._log_toggle_btn = tk.Button(
            log_hdr, text="▶  Logs",
            font=("Segoe UI", 9), fg=TEXT_MUTED, bg=APP_BG,
            activebackground=APP_BG, activeforeground=ACCENT,
            relief="flat", cursor="hand2", anchor="w", bd=0,
            command=self._toggle_logs,
        )
        self._log_toggle_btn.pack(side="left")

        self._log_clear_btn = tk.Button(
            log_hdr, text="Clear",
            font=("Segoe UI", 8), fg=TEXT_MUTED, bg=APP_BG,
            activebackground=APP_BG, activeforeground=ACCENT,
            relief="flat", cursor="hand2", bd=0,
            command=self._clear_logs,
        )
        self._log_clear_btn.pack(side="right")
        self._log_clear_btn.pack_forget()  # hidden until expanded

        self._log_body = tk.Frame(log_outer, bg=APP_BG)
        # body not packed until expanded

        log_scroll = ttk.Scrollbar(self._log_body, orient="vertical")
        self._log_text = tk.Text(
            self._log_body,
            font=("Consolas", 8),
            fg=TEXT_PRIMARY, bg="#f8f9fa",
            relief="flat", bd=0, highlightthickness=1,
            highlightbackground=BORDER,
            wrap="word", height=10,
            state="disabled",
            yscrollcommand=log_scroll.set,
        )
        log_scroll.config(command=self._log_text.yview)
        log_scroll.pack(side="right", fill="y")
        self._log_text.pack(fill="both", expand=True)

        self._populate()
        win.update_idletasks()
        # centre on screen
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        ww = win.winfo_reqwidth()
        wh = win.winfo_reqheight()
        win.geometry(f"+{(sw - ww)//2}+{(sh - wh)//2}")
        win.deiconify()
        
        # Remove minimize button after window is shown
        self._remove_minimize_button()

    def _apply_style(self):
        style = ttk.Style(self._root)
        style.theme_use("clam")
        style.configure("Vertical.TScrollbar",
                        background=PANEL_BG, troughcolor=APP_BG,
                        arrowcolor=TEXT_MUTED, bordercolor=BORDER)
        style.configure("Switch.TCheckbutton",
                        background=APP_BG, foreground=TEXT_PRIMARY,
                        font=("Segoe UI", 9))
        style.map("Switch.TCheckbutton",
                  background=[("active", APP_BG)],
                  foreground=[("active", ACCENT)])

        style.configure("Window.Treeview",
                        background=PANEL_BG,
                        foreground=TEXT_PRIMARY,
                        fieldbackground=PANEL_BG,
                        borderwidth=0,
                        relief="flat",
                        font=("Segoe UI", 10),
                        rowheight=28)
        style.configure("Window.Treeview.Item", padding=(10, 2))
        style.map("Window.Treeview",
                  background=[("selected", ACCENT)],
                  foreground=[("selected", "#ffffff")])

    def _set_resize_btn_state(self, enabled: bool):
        if enabled:
            self._resize_btn.config(
                state="normal",
                bg=ACCENT, fg="#ffffff",
                activebackground=ACCENT_HOVER, activeforeground="#ffffff",
                cursor="hand2",
            )
        else:
            self._resize_btn.config(
                state="disabled",
                bg=BTN_DISABLED_BG, fg=BTN_DISABLED_FG,
                activebackground=BTN_DISABLED_BG, activeforeground=BTN_DISABLED_FG,
                cursor="",
            )

    def _populate(self):
        for iid in self._tree.get_children():
            self._tree.delete(iid)
        self._hwnd_map.clear()
        self._icon_cache.clear()
        self._size_var.set("- select a window -")
        self._status_var.set("")
        self._set_resize_btn_state(False)

        self._tree.tag_configure("even", background=PANEL_BG)
        self._tree.tag_configure("odd",  background=ROW_ALT)

        windows = _enum_windows()
        for i, (hwnd, title) in enumerate(windows):
            display = title[:80] + ("…" if len(title) > 80 else "")
            tag = "odd" if i % 2 else "even"

            photo: ImageTk.PhotoImage | None = None
            icon_img = _get_window_icon(hwnd, 20)
            if icon_img:
                try:
                    photo = ImageTk.PhotoImage(icon_img)
                    self._icon_cache.append(photo)
                except Exception:
                    pass

            iid = self._tree.insert("", tk.END, text=f"  {display}",
                                    image=photo if photo else "", tags=(tag,))
            self._hwnd_map[iid] = hwnd

        self._count_var.set(f"{len(windows)} windows")

    def _selected_hwnds(self) -> list[int]:
        return [self._hwnd_map[iid] for iid in self._tree.selection()
                if iid in self._hwnd_map]

    def _on_select(self, _event=None):
        hwnds = self._selected_hwnds()
        if not hwnds:
            self._size_var.set("- select a window -")
            self._set_resize_btn_state(False)
            return

        self._set_resize_btn_state(True)

        if len(hwnds) > 1:
            self._size_var.set(f"{len(hwnds)} windows selected")
            self._status_lbl.config(fg=TEXT_PRIMARY)
            return

        hwnd = hwnds[0]
        w, h = _get_client_size(hwnd)
        ncl, nct, ncr, ncb = _get_nonclient_margins(hwnd)
        color = SUCCESS if (w == TARGET_W and h == TARGET_H) else WARNING
        parts = []
        if nct:
            parts.append(f"titlebar {nct}px")
        border = max(ncl, ncr, ncb)
        if border:
            parts.append(f"border {border}px")
        note = f"  ({', '.join(parts)})" if parts else ""
        self._size_var.set(f"{w} × {h}{note}")
        self._status_lbl.config(fg=color)

    def _toggle_logs(self):
        if self._log_expanded:
            self._log_body.pack_forget()
            self._log_clear_btn.pack_forget()
            self._log_toggle_btn.config(text="▶  Logs")
            self._log_expanded = False
        else:
            self._log_body.pack(fill="both", expand=True, pady=(4, 0))
            self._log_clear_btn.pack(side="right")
            self._log_toggle_btn.config(text="▼  Logs")
            self._log_expanded = True

    def _append_log(self, text: str):
        self._log_text.config(state="normal")
        self._log_text.insert("end", text + "\n")
        self._log_text.see("end")
        self._log_text.config(state="disabled")

    def _clear_logs(self):
        self._log_text.config(state="normal")
        self._log_text.delete("1.0", "end")
        self._log_text.config(state="disabled")

    def _do_resize(self):
        hwnds = self._selected_hwnds()
        if not hwnds:
            messagebox.showwarning("No selection", "Please select a window first.")
            return
        move = self._move_var.get()
        results = [_resize_window(hwnd, move, self._append_log) for hwnd in hwnds]
        if len(results) == 1:
            msg = results[0]
        else:
            ok = sum(1 for r in results if "✓" in r)
            msg = f"{ok}/{len(results)} windows resized ✓" if ok == len(results) \
                  else f"{ok}/{len(results)} succeeded"
        self._status_var.set(msg)
        self._status_lbl.config(fg=SUCCESS if "✓" in msg else WARNING)
        self._on_select()   # refresh size display

    def _remove_minimize_button(self):
        """Remove the minimize button from the window using Win32 API."""
        if not self._win:
            return
        try:
            # Get the window handle
            hwnd = self._win.winfo_id()
            
            # Get current window style
            style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
            
            # Remove the minimize button (WS_MINIMIZEBOX)
            WS_MINIMIZEBOX = 0x00020000
            new_style = style & ~WS_MINIMIZEBOX
            
            # Apply the new style
            win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, new_style)
            
            # Force the window frame to be redrawn
            SWP_FRAMECHANGED = 0x0020
            SWP_NOMOVE = 0x0002
            SWP_NOSIZE = 0x0001
            SWP_NOZORDER = 0x0004
            win32gui.SetWindowPos(hwnd, 0, 0, 0, 0, 0, 
                                  SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER)
        except Exception as e:
            # Silently ignore errors - this is a cosmetic feature
            pass

    def _on_close_to_tray(self):
        """Hide the window to system tray instead of exiting."""
        if self._win:
            self._win.withdraw()


# ── Application entry point ───────────────────────────────────────────────────
class App:
    def __init__(self):
        # Hidden root window - owns the event loop
        self._root = tk.Tk()
        self._root.withdraw()
        self._root.title("1080p Window Resizer")
        # Set icon on the hidden root so every child Toplevel inherits it
        self._tk_icon = ImageTk.PhotoImage(_load_app_icon(32))
        self._root.iconphoto(True, self._tk_icon)

        self._tray: pystray.Icon | None = None
        self._selector = SelectorWindow(self._root)
        self._root.protocol("WM_DELETE_WINDOW", self._on_exit)

    # ── Tray setup ────────────────────────────────────────────────────────────
    def _build_tray(self):
        menu = pystray.Menu(
            pystray.MenuItem("Show Window",      self._on_show_window, default=True),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Exit",             self._on_exit),
        )
        self._tray = pystray.Icon(
            "StreamResizer",
            _make_tray_image(),
            "1080p Window Resizer\nClick to show window",
            menu,
        )

    def _on_show_window(self, _icon=None, _item=None):
        self._selector.show()

    def _on_exit(self, _icon=None, _item=None):
        if self._tray:
            self._tray.stop()
        self._root.after(0, self._root.quit)

    # ── Run ───────────────────────────────────────────────────────────────────
    def run(self):
        self._build_tray()

        # pystray must run on its own thread; tkinter owns the main thread
        # WHY: Both Tkinter (mainloop) and pystray (icon.run) are blocking calls.
        # We can't run them both on the same thread. Since Tkinter MUST be on the main thread
        # (on macOS/Windows mostly), we push the tray icon to a background thread.
        tray_thread = threading.Thread(target=self._tray.run, daemon=True)
        tray_thread.start()

        # Open the selector window immediately on launch
        self._selector.show()

        self._root.mainloop()


if __name__ == "__main__":
    App().run()
