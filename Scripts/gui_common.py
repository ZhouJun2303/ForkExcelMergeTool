# -*- coding: utf-8 -*-
"""
GUI 公共组件与样式：日志输出到文件并更新状态栏、颜色图例、用系统默认程序打开文件、合并/对比窗口统一样式。
只做一件事：为 Merge/Diff 窗口提供统一的写日志、界面小部件和样式，不包含业务逻辑。
"""

import os
import math
import subprocess
import sys
import threading
import tkinter as tk
import time
from tkinter import ttk, messagebox
from datetime import datetime

from log_util import log_path, log
from update_manager import (
    UpdateError,
    cached_update_info,
    check_for_update,
    download_update,
    get_current_executable,
    launch_update_script,
    make_update_script,
    remember_update_check,
    should_auto_check_update,
)


def gui_log_path():
    """GUI 下使用的日志文件路径（与主入口一致）。"""
    return log_path()


def gui_log(msg, log_widget=None, is_error=False):
    """
    写一行到日志文件，并可选更新 GUI 上的日志控件（如状态栏 StringVar）。
    log_widget: 若为 tk.StringVar 则 set 为单行；若为 Text 则 insert 并滚动到底。
    """
    ts = datetime.now().strftime("%H:%M:%S")
    prefix = "[ERROR] " if is_error else ""
    line = "%s %s%s" % (ts, prefix, msg)
    try:
        with open(log_path(), "a", encoding="utf-8") as f:
            f.write(datetime.now().strftime("%Y-%m-%d ") + line + "\n")
    except Exception:
        pass
    if log_widget:
        try:
            if isinstance(log_widget, tk.StringVar):
                log_widget.set(line.strip())
            else:
                log_widget.insert(tk.END, line + "\n")
                log_widget.see(tk.END)
            if hasattr(log_widget, "update_idletasks"):
                log_widget.update_idletasks()
        except Exception:
            pass


UI = {
    "bg": "#F4F7FB",
    "panel": "#FFFFFF",
    "panel_alt": "#F8FAFC",
    "border": "#D8E0EA",
    "text": "#172033",
    "muted": "#64748B",
    "primary": "#2563EB",
    "primary_active": "#1D4ED8",
    "success": "#15803D",
    "success_bg": "#DCFCE7",
    "warning": "#B45309",
    "warning_bg": "#FEF3C7",
    "danger": "#B91C1C",
    "danger_bg": "#FEE2E2",
    "deleted": "#64748B",
    "deleted_bg": "#EEF2F7",
    "row_alt": "#F8FAFC",
}


def resource_path(*parts):
    """返回源码/ PyInstaller 运行时均可访问的资源路径。"""
    env_root = os.environ.get("EXCEL_MERGE_FORK_RESOURCE_ROOT")
    if env_root:
        base = os.path.abspath(env_root)
    elif getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *parts)


def ui_font(size=9, weight="normal"):
    """统一 UI 字体，优先使用 Windows 中文界面更自然的微软雅黑。"""
    return ("Microsoft YaHei UI", size, weight)


def _scaled(value, size):
    return max(0, int(round(value * size / 16.0)))


def _line(img, points, color, size=16, thickness=1):
    """在 PhotoImage 上画一条简单折线，足够支撑小图标。"""
    last = None
    t = max(1, _scaled(thickness, size))
    for raw_x, raw_y in points:
        x = _scaled(raw_x, size)
        y = _scaled(raw_y, size)
        if last is not None:
            lx, ly = last
            dx = x - lx
            dy = y - ly
            steps = max(abs(dx), abs(dy), 1)
            for i in range(steps + 1):
                px = int(round(lx + dx * i / float(steps)))
                py = int(round(ly + dy * i / float(steps)))
                img.put(color, to=(px, py, min(size, px + t), min(size, py + t)))
        last = (x, y)


def _icon_asset_path(name, size):
    asset_names = {
        "download": "update_available",
        "refresh": "update_checking",
    }
    asset_name = asset_names.get(name, name)
    preferred_size = 32 if size >= 24 and asset_name == "app" else 16
    path = resource_path("Assets", "icons", "%s_%d.png" % (asset_name, preferred_size))
    if os.path.isfile(path):
        return path
    fallback = resource_path("Assets", "icons", "%s_16.png" % asset_name)
    if os.path.isfile(fallback):
        return fallback
    return None


def get_ui_icon(root, name, size=16, color=None):
    """创建并缓存一组轻量 Tk 图标，避免额外图片依赖。"""
    cache = getattr(root, "_excel_merge_ui_icons", None)
    if cache is None:
        cache = {}
        root._excel_merge_ui_icons = cache
    fg = color or UI["primary"]
    key = (name, size, fg)
    if key in cache:
        return cache[key]

    asset_path = _icon_asset_path(name, size)
    if asset_path and color is None:
        try:
            img = tk.PhotoImage(master=root, file=asset_path)
            cache[key] = img
            return img
        except Exception:
            pass

    img = tk.PhotoImage(master=root, width=size, height=size)

    def rect(x1, y1, x2, y2, fill):
        sx1, sy1, sx2, sy2 = (_scaled(x1, size), _scaled(y1, size), _scaled(x2, size), _scaled(y2, size))
        if sx2 <= sx1:
            sx2 = sx1 + 1
        if sy2 <= sy1:
            sy2 = sy1 + 1
        img.put(fill, to=(sx1, sy1, min(size, sx2), min(size, sy2)))

    def line(points, fill=fg, thickness=1):
        _line(img, points, fill, size=size, thickness=thickness)

    muted = UI["muted"]
    green = UI["success"]
    yellow = UI["warning"]
    red = UI["danger"]
    pale_blue = "#DBEAFE"
    pale_gray = UI["deleted_bg"]

    if name == "app":
        rect(2, 1, 14, 15, pale_blue)
        rect(3, 2, 13, 14, UI["panel"])
        rect(3, 2, 13, 5, green)
        rect(5, 7, 8, 8, green)
        rect(9, 7, 12, 8, green)
        rect(5, 10, 8, 11, green)
        rect(9, 10, 12, 11, green)
        line([(3, 5), (13, 5)], muted)
    elif name in ("update", "download"):
        rect(3, 12, 13, 14, pale_blue)
        rect(4, 13, 12, 14, fg)
        rect(7, 2, 9, 9, fg)
        line([(4, 7), (8, 11), (12, 7)], fg, 2)
    elif name == "refresh":
        line([(4, 5), (6, 3), (10, 3), (12, 5), (12, 7)], fg, 2)
        line([(12, 11), (10, 13), (6, 13), (4, 11), (4, 9)], fg, 2)
        line([(10, 5), (12, 7), (14, 5)], fg, 1)
        line([(6, 11), (4, 9), (2, 11)], fg, 1)
    elif name == "open":
        rect(4, 2, 11, 14, pale_gray)
        rect(5, 3, 10, 13, UI["panel"])
        rect(10, 3, 13, 6, pale_gray)
        line([(10, 3), (13, 6), (10, 6), (10, 3)], muted)
        rect(6, 8, 11, 9, fg)
        rect(6, 10, 10, 11, fg)
    elif name == "copy":
        rect(5, 2, 13, 10, pale_gray)
        line([(5, 2), (13, 2), (13, 10), (5, 10), (5, 2)], muted)
        rect(3, 6, 11, 14, UI["panel"])
        line([(3, 6), (11, 6), (11, 14), (3, 14), (3, 6)], fg)
        rect(5, 9, 9, 10, fg)
        rect(5, 11, 8, 12, fg)
    elif name == "folder":
        rect(2, 5, 7, 7, "#FDE68A")
        rect(2, 7, 14, 13, "#FEF3C7")
        rect(3, 8, 13, 12, yellow)
    elif name == "merge":
        rect(3, 2, 6, 5, fg)
        rect(3, 11, 6, 14, fg)
        rect(11, 6, 14, 9, green)
        line([(5, 5), (5, 8), (11, 8)], fg, 2)
        line([(5, 11), (5, 8)], fg, 2)
    elif name == "check":
        line([(3, 8), (7, 12), (13, 4)], green, 2)
    elif name == "cancel":
        line([(4, 4), (12, 12)], red, 2)
        line([(12, 4), (4, 12)], red, 2)
    elif name == "backup":
        rect(3, 2, 13, 14, pale_gray)
        rect(4, 3, 12, 7, fg)
        rect(6, 10, 10, 13, UI["panel"])
        rect(9, 3, 11, 6, UI["panel"])
    elif name == "swap":
        line([(3, 5), (12, 5)], fg, 2)
        line([(10, 3), (13, 5), (10, 7)], fg, 1)
        line([(13, 11), (4, 11)], green, 2)
        line([(6, 9), (3, 11), (6, 13)], green, 1)
    elif name == "detail":
        rect(4, 4, 10, 10, UI["panel"])
        line([(4, 4), (10, 4), (10, 10), (4, 10), (4, 4)], fg, 1)
        line([(9, 9), (13, 13)], fg, 2)
    elif name == "prev":
        line([(11, 3), (5, 8), (11, 13)], fg, 2)
    elif name == "next":
        line([(5, 3), (11, 8), (5, 13)], fg, 2)
    else:
        rect(4, 4, 12, 12, fg)

    cache[key] = img
    return img


def configure_button_icon(root, button, icon_name, size=16, color=None):
    """给 ttk.Button 配置小图标，并缓存引用避免被 Tk 回收。"""
    try:
        img = get_ui_icon(root, icon_name, size=size, color=color)
        button.configure(image=img, compound=tk.LEFT)
        button._excel_merge_icon = img
    except Exception:
        pass
    return button


def make_icon_button(parent, root, text, icon_name, command=None, style="Secondary.TButton", **kwargs):
    """创建带统一小图标的按钮。"""
    button = ttk.Button(parent, text=text, command=command, style=style, **kwargs)
    configure_button_icon(root, button, icon_name)
    return button


def make_header_icon(parent, root, style_name="Icon.TLabel"):
    """标题区应用图标。"""
    img = get_ui_icon(root, "app", size=32)
    label = ttk.Label(parent, image=img, style=style_name)
    label._excel_merge_icon = img
    return label


def make_update_card(parent, root, include_button=True):
    """创建现代化更新状态卡片，返回 controller 需要绑定的控件。"""
    card = ttk.Frame(parent, padding=(10, 8), style="UpdateCard.TFrame")
    title_row = ttk.Frame(card, style="UpdateCard.TFrame")
    title_row.pack(fill=tk.X)
    title_text = "程序更新" if include_button else "更新状态"
    ttk.Label(title_row, text=title_text, style="UpdateTitle.TLabel").pack(side=tk.LEFT)
    btn_update = None
    if include_button:
        btn_update = ttk.Button(title_row, text="手动检查更新", style="Tool.TButton")
        btn_update.pack(side=tk.RIGHT)
    state_var = tk.StringVar(root, value="")
    state_row = ttk.Frame(card, style="UpdateCard.TFrame")
    state_row.pack(fill=tk.X, pady=(5, 0))
    state_icon = ttk.Label(state_row, image=get_ui_icon(root, "update", size=16), style="UpdateIcon.TLabel")
    state_icon._excel_merge_icon = get_ui_icon(root, "update", size=16)
    state_icon.pack(side=tk.LEFT, padx=(0, 6))
    state_label = ttk.Label(state_row, textvariable=state_var, style="UpdateIdle.TLabel")
    state_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

    progress_row = ttk.Frame(card, style="UpdateCard.TFrame")
    progress_var = tk.IntVar(root, value=0)
    progress_label = ttk.Label(progress_row, text="", style="UpdateIdle.TLabel")
    progress_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 8))
    progress_bar = ttk.Progressbar(progress_row, variable=progress_var, maximum=100, length=190)
    progress_bar.grid(row=0, column=1, sticky=tk.EW)
    progress_row.columnconfigure(1, weight=1)
    return card, btn_update, state_var, state_label, state_icon, progress_row, progress_var, progress_label, progress_bar


def apply_app_icon(root):
    """设置窗口/任务栏图标。"""
    ico_path = resource_path("Assets", "ExcelMergeFork.ico")
    if os.path.isfile(ico_path):
        try:
            root.iconbitmap(ico_path)
            return True
        except Exception:
            pass
    try:
        icon = get_ui_icon(root, "app", size=32)
        root.iconphoto(True, icon)
        return True
    except Exception:
        return False


def make_color_legend(parent, items, bg=None):
    """
    在 parent 下生成一行颜色图例。
    items: [(颜色 hex 或 None, "说明文字"), ...]
    """
    bg = bg or UI["panel"]
    frame = tk.Frame(parent, bg=bg)
    for i, (color, text) in enumerate(items):
        if i > 0:
            tk.Label(frame, text="  ", bg=bg, font=ui_font(8)).pack(side=tk.LEFT)
        if color:
            patch = tk.Frame(frame, width=14, height=14, bg=color, relief=tk.SOLID, borderwidth=1)
            patch.pack(side=tk.LEFT, padx=(0, 4))
            patch.pack_propagate(False)
        tk.Label(frame, text=text, bg=bg, font=ui_font(8), fg=UI["muted"]).pack(side=tk.LEFT)
    return frame


def make_badge(parent, text, tone="neutral"):
    """创建一个小状态徽标。tone: primary|success|warning|danger|neutral。"""
    colors = {
        "primary": (UI["primary"], "#DBEAFE"),
        "success": (UI["success"], UI["success_bg"]),
        "warning": (UI["warning"], UI["warning_bg"]),
        "danger": (UI["danger"], UI["danger_bg"]),
        "neutral": (UI["muted"], UI["deleted_bg"]),
    }
    fg, bg = colors.get(tone, colors["neutral"])
    return tk.Label(
        parent,
        text=text,
        fg=fg,
        bg=bg,
        font=ui_font(8, "bold"),
        padx=8,
        pady=2,
        relief=tk.FLAT,
    )


def make_separator(parent):
    line = tk.Frame(parent, bg=UI["border"], height=1)
    line.pack_propagate(False)
    return line


class GlobalBusyIndicator:
    """Small overlay spinner shared by all GUI windows."""

    def __init__(self, root):
        self.root = root
        self._tokens = {}
        self._hide_after_id = None
        self._spin_after_id = None
        self._spin_step = 0
        self._visible_since = 0.0
        self._min_visible_ms = 260
        self._frame = tk.Frame(root, bg=UI["border"], borderwidth=1, relief=tk.SOLID)
        inner = ttk.Frame(self._frame, padding=(10, 6), style="Loading.TFrame")
        inner.pack(fill=tk.BOTH, expand=True)
        self._spinner = tk.Canvas(
            inner,
            width=22,
            height=22,
            bg=UI["panel"],
            highlightthickness=0,
            borderwidth=0,
        )
        self._spinner.pack(side=tk.LEFT)
        self._label_var = tk.StringVar(root, value="处理中...")
        ttk.Label(inner, textvariable=self._label_var, style="Loading.TLabel").pack(side=tk.LEFT, padx=(7, 0))

    def pulse(self, text="处理中...", duration_ms=650):
        token = ("click", id(object()), time.monotonic())
        self.show(token, text)
        try:
            self.root.after(duration_ms, lambda t=token: self.hide(t))
        except tk.TclError:
            pass
        return token

    def set_task(self, key, busy, text="处理中..."):
        token = ("task", key)
        if busy:
            self.show(token, text)
        else:
            self.hide(token)

    def show(self, token, text="处理中..."):
        if self._hide_after_id is not None:
            try:
                self.root.after_cancel(self._hide_after_id)
            except tk.TclError:
                pass
            self._hide_after_id = None
        self._tokens[token] = text or "处理中..."
        self._label_var.set(self._tokens[token])
        if not self._frame.winfo_ismapped():
            self._visible_since = time.monotonic()
            try:
                self._frame.place(relx=1.0, x=-14, y=12, anchor=tk.NE)
            except tk.TclError:
                return
        try:
            self._frame.lift()
            self.root.update_idletasks()
            self._start_spin()
        except tk.TclError:
            pass

    def hide(self, token=None):
        if token is None:
            self._tokens.clear()
        else:
            self._tokens.pop(token, None)
        if self._tokens:
            latest_text = list(self._tokens.values())[-1]
            self._label_var.set(latest_text)
            return
        elapsed_ms = int((time.monotonic() - self._visible_since) * 1000)
        delay = max(0, self._min_visible_ms - elapsed_ms)
        if delay:
            try:
                self._hide_after_id = self.root.after(delay, self._hide_now)
            except tk.TclError:
                pass
        else:
            self._hide_now()

    def _start_spin(self):
        if self._spin_after_id is None:
            self._draw_spinner()

    def _draw_spinner(self):
        if not self._frame.winfo_ismapped():
            self._spin_after_id = None
            return
        try:
            self._spinner.delete("all")
            colors = ["#2563EB", "#3B82F6", "#60A5FA", "#93C5FD", "#BFDBFE", "#DBEAFE", "#E0E7FF", "#CBD5E1"]
            cx = cy = 11
            radius = 7
            for i in range(8):
                angle = (self._spin_step + i) * math.pi / 4.0
                x = cx + math.cos(angle) * radius
                y = cy + math.sin(angle) * radius
                self._spinner.create_oval(x - 2, y - 2, x + 2, y + 2, fill=colors[i], outline="")
            self._spin_step = (self._spin_step + 1) % 8
            self._spin_after_id = self.root.after(80, self._draw_spinner)
        except tk.TclError:
            self._spin_after_id = None

    def _hide_now(self):
        self._hide_after_id = None
        if self._tokens:
            return
        try:
            self._frame.place_forget()
            if self._spin_after_id is not None:
                self.root.after_cancel(self._spin_after_id)
            self._spin_after_id = None
            self._spinner.delete("all")
        except tk.TclError:
            pass


def get_global_busy_indicator(root):
    indicator = getattr(root, "_excel_merge_global_busy", None)
    if indicator is None:
        indicator = GlobalBusyIndicator(root)
        root._excel_merge_global_busy = indicator
    return indicator


def set_global_busy(root, key, busy, text=None):
    try:
        get_global_busy_indicator(root).set_task(key, busy, text or "处理中...")
    except Exception:
        pass


def run_loading_task(root, key, text, worker_func, success_func, error_func=None):
    """
    Run a worker in the background while showing the shared loading indicator.

    worker_func runs off the Tk thread. success_func/error_func run back on the
    Tk thread after the loading token is released.
    """
    token = ("loading_task", key, id(worker_func), time.monotonic())
    try:
        get_global_busy_indicator(root).show(token, text or "处理中...")
    except Exception:
        pass

    def finish(result=None, err=None):
        try:
            get_global_busy_indicator(root).hide(token)
        except Exception:
            pass
        if err is None:
            if success_func is not None:
                success_func(result)
            return
        if error_func is not None:
            error_func(err)
        else:
            messagebox.showerror("错误", str(err))

    def worker():
        result = None
        err = None
        try:
            result = worker_func()
        except Exception as exc:
            err = exc
        try:
            root.after(0, lambda result=result, err=err: finish(result, err))
        except Exception:
            pass

    threading.Thread(target=worker, daemon=True).start()
    return token


def install_global_button_loading(root):
    """Show a short global spinner whenever a Tk/ttk button is pressed."""
    try:
        if root.tk.getvar("excel_merge_loading_bind_installed") == "1":
            return
    except tk.TclError:
        pass

    def _button_finished(widget, expected_token=None):
        try:
            token = getattr(widget, "_excel_merge_loading_token", None)
            if token is None or (expected_token is not None and token != expected_token):
                return
            widget._excel_merge_loading_token = None
            indicator = getattr(widget, "_excel_merge_loading_indicator", None)
            if indicator is not None:
                indicator.hide(token)
        except Exception:
            pass

    def _button_pressed(event):
        widget = getattr(event, "widget", None)
        if widget is None:
            return
        try:
            if hasattr(widget, "state") and "disabled" in widget.state():
                return
            if str(widget.cget("state")) == tk.DISABLED:
                return
        except Exception:
            pass
        try:
            top = widget.winfo_toplevel()
            indicator = get_global_busy_indicator(top)
            token = ("button", id(widget), time.monotonic())
            widget._excel_merge_loading_token = token
            widget._excel_merge_loading_indicator = indicator
            indicator.show(token, "处理中...")
            top.after(30000, lambda w=widget, t=token: _button_finished(w, t))
        except Exception:
            pass

    def _button_released(event):
        widget = getattr(event, "widget", None)
        if widget is not None:
            _button_finished(widget)

    root.bind_class("TButton", "<ButtonPress-1>", _button_pressed, add="+")
    root.bind_class("Button", "<ButtonPress-1>", _button_pressed, add="+")
    root.bind_class("TButton", "<ButtonRelease-1>", _button_released, add="+")
    root.bind_class("Button", "<ButtonRelease-1>", _button_released, add="+")
    try:
        root.tk.setvar("excel_merge_loading_bind_installed", "1")
    except tk.TclError:
        pass


class ToolTip:
    """轻量 tooltip，避免引入额外依赖。"""

    def __init__(self, widget, text, delay=450):
        self.widget = widget
        self.text = text
        self.delay = delay
        self._after_id = None
        self._tip = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _event=None):
        self._cancel()
        self._after_id = self.widget.after(self.delay, self._show)

    def _cancel(self):
        if self._after_id is not None:
            try:
                self.widget.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

    def _show(self):
        if self._tip is not None or not self.text:
            return
        try:
            x = self.widget.winfo_rootx() + 12
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        except Exception:
            return
        tip = tk.Toplevel(self.widget)
        tip.wm_overrideredirect(True)
        tip.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(
            tip,
            text=self.text,
            justify=tk.LEFT,
            bg="#111827",
            fg="#F9FAFB",
            font=ui_font(8),
            padx=8,
            pady=5,
            wraplength=360,
        )
        label.pack()
        self._tip = tip

    def _hide(self, _event=None):
        self._cancel()
        if self._tip is not None:
            try:
                self._tip.destroy()
            except Exception:
                pass
            self._tip = None


def open_excel_file(path):
    """用系统默认程序打开 Excel 文件；路径不存在或打开失败返回 False。"""
    if not path or not os.path.isfile(path):
        return False
    try:
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.run(["open", path], check=False)
        else:
            subprocess.run(["xdg-open", path], check=False)
        return True
    except Exception:
        return False


def open_containing_folder(path, select_file=True):
    """
    打开文件所在文件夹。
    - Windows: 优先 explorer /select 选中文件；失败则打开目录
    - macOS: open -R
    - Linux: xdg-open 目录
    """
    if not path:
        return False
    try:
        abs_path = os.path.abspath(path)
        folder = abs_path if os.path.isdir(abs_path) else os.path.dirname(abs_path)
        if not folder:
            return False
        if sys.platform == "win32":
            if select_file and os.path.isfile(abs_path):
                try:
                    subprocess.run(["explorer", "/select,", abs_path], check=False)
                    return True
                except Exception:
                    pass
            os.startfile(folder)
            return True
        if sys.platform == "darwin":
            if select_file and os.path.exists(abs_path):
                subprocess.run(["open", "-R", abs_path], check=False)
            else:
                subprocess.run(["open", folder], check=False)
            return True
        subprocess.run(["xdg-open", folder], check=False)
        return True
    except Exception:
        return False


def setup_merge_styles(root):
    """为合并/对比窗口配置统一现代桌面样式。"""
    style = ttk.Style(root)
    try:
        style.theme_use("vista")
    except tk.TclError:
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

    bg = UI["bg"]
    panel = UI["panel"]
    fg = UI["text"]
    muted = UI["muted"]

    style.configure(".", font=ui_font(9))
    style.configure("TFrame", background=bg)
    style.configure("App.TFrame", background=bg)
    style.configure("Panel.TFrame", background=panel, relief=tk.FLAT)
    style.configure("Card.TFrame", background=UI["panel_alt"], relief=tk.FLAT)
    style.configure("UpdateCard.TFrame", background=UI["panel_alt"], relief=tk.FLAT)
    style.configure("Loading.TFrame", background=panel, relief=tk.FLAT)
    style.configure("Toolbar.TFrame", background=panel)
    style.configure("BottomBar.TFrame", background=panel)
    style.configure("TLabel", background=bg, foreground=fg, font=ui_font(9))
    style.configure("Icon.TLabel", background=bg)
    style.configure("PanelIcon.TLabel", background=panel)
    style.configure("UpdateIcon.TLabel", background=UI["panel_alt"])
    style.configure("Panel.TLabel", background=panel, foreground=fg, font=ui_font(9))
    style.configure("Card.TLabel", background=UI["panel_alt"], foreground=fg, font=ui_font(9))
    style.configure("CardMuted.TLabel", background=UI["panel_alt"], foreground=muted, font=ui_font(8))
    style.configure("CardSection.TLabel", background=UI["panel_alt"], foreground=fg, font=ui_font(10, "bold"))
    style.configure("UpdateTitle.TLabel", background=UI["panel_alt"], foreground=fg, font=ui_font(9, "bold"))
    style.configure("UpdateIdle.TLabel", background=UI["panel_alt"], foreground=muted, font=ui_font(8))
    style.configure("UpdateChecking.TLabel", background=UI["panel_alt"], foreground=UI["primary"], font=ui_font(8, "bold"))
    style.configure("UpdateReady.TLabel", background=UI["panel_alt"], foreground=UI["success"], font=ui_font(8, "bold"))
    style.configure("UpdateError.TLabel", background=UI["panel_alt"], foreground=UI["danger"], font=ui_font(8, "bold"))
    style.configure("Muted.TLabel", background=panel, foreground=muted, font=ui_font(8))
    style.configure("Loading.TLabel", background=panel, foreground=fg, font=ui_font(9, "bold"))
    style.configure("Title.TLabel", background=bg, foreground=fg, font=ui_font(16, "bold"))
    style.configure("PanelTitle.TLabel", background=panel, foreground=fg, font=ui_font(16, "bold"))
    style.configure("Subtitle.TLabel", background=bg, foreground=muted, font=ui_font(9))
    style.configure("Section.TLabel", background=panel, foreground=fg, font=ui_font(10, "bold"))
    style.configure("TLabelframe", background=panel, borderwidth=1, relief=tk.SOLID)
    style.configure("TLabelframe.Label", background=panel, foreground=fg, font=ui_font(9, "bold"))
    style.configure("TButton", font=ui_font(9), padding=(12, 6))
    style.configure("Secondary.TButton", font=ui_font(9), padding=(10, 6))
    style.configure("Tiny.TButton", font=ui_font(8), padding=(7, 3))
    style.configure("Accent.TButton", font=ui_font(9, "bold"), padding=(14, 7))
    style.configure("Tool.TButton", font=ui_font(9), padding=(9, 5))
    style.map(
        "TButton",
        foreground=[("disabled", "#64748B"), ("!disabled", fg)],
        background=[("active", "#E2E8F0"), ("disabled", "#F1F5F9"), ("!disabled", "#FFFFFF")],
    )
    style.map(
        "Secondary.TButton",
        foreground=[("disabled", "#64748B"), ("!disabled", fg)],
        background=[("active", "#E2E8F0"), ("disabled", "#F1F5F9"), ("!disabled", "#FFFFFF")],
    )
    style.map(
        "Tiny.TButton",
        foreground=[("disabled", "#64748B"), ("!disabled", fg)],
        background=[("active", "#E2E8F0"), ("disabled", "#F1F5F9"), ("!disabled", "#FFFFFF")],
    )
    style.map(
        "Accent.TButton",
        foreground=[("disabled", "#64748B"), ("!disabled", fg)],
        background=[("active", "#DBEAFE"), ("disabled", "#F1F5F9"), ("!disabled", "#FFFFFF")],
    )
    style.map(
        "Tool.TButton",
        foreground=[("disabled", "#64748B"), ("!disabled", fg)],
        background=[("active", "#E2E8F0"), ("disabled", "#F1F5F9"), ("!disabled", "#FFFFFF")],
    )
    style.configure("TCheckbutton", background=panel, foreground=fg, font=ui_font(9), padding=(2, 4))
    style.configure("TRadiobutton", background=panel, foreground=fg, font=ui_font(9), padding=(2, 4))
    style.configure("TCombobox", font=ui_font(9), padding=(8, 4))
    style.configure("TEntry", font=ui_font(9), padding=(8, 4))
    style.configure(
        "Treeview",
        font=ui_font(9),
        rowheight=28,
        fieldbackground=panel,
        background=panel,
        foreground=fg,
        borderwidth=0,
    )
    style.configure("Treeview.Heading", font=ui_font(9, "bold"), background=UI["panel_alt"], foreground=fg)
    style.map(
        "Treeview",
        background=[("selected", "#DBEAFE")],
        foreground=[("selected", fg)],
    )
    root.configure(bg=bg)


class UpdateButtonController:
    """把 GitHub Release 更新检查接到一个 Tk 按钮上。"""

    def __init__(self, root, button, status_var=None, on_quit=None, compact=False):
        self.root = root
        self.button = button
        self.status_var = status_var
        self.on_quit = on_quit
        self.compact = compact
        self.default_text = "手动检查更新"
        self.checking_text = "检查中..."
        self.available_text = None
        self.downloading_text = "下载中..."
        self.progress_container = None
        self.progress_var = None
        self.progress_bar = None
        self.progress_label = None
        self.state_var = None
        self.state_label = None
        self.state_icon = None
        self.info = None
        self.checking = False
        self.installing = False
        self.button.config(text=self.default_text, command=self.on_click)
        configure_button_icon(self.root, self.button, "update")
        self._set_state_text("尚未检查更新", "idle")
        self._apply_cached_info()

    def bind_progress_widgets(self, progress_bar, progress_var, progress_label=None, progress_container=None):
        self.progress_container = progress_container
        self.progress_bar = progress_bar
        self.progress_var = progress_var
        self.progress_label = progress_label
        self._set_progress_visible(False)

    def bind_state_widget(self, state_var, state_label=None, state_icon=None):
        self.state_var = state_var
        self.state_label = state_label
        self.state_icon = state_icon
        self._apply_cached_info()
        if not self.info:
            self._set_state_text("尚未检查更新", "idle")

    def start_background_check(self):
        cached = cached_update_info()
        if cached and cached.get("available"):
            self._set_check_result(cached, silent=True, from_cache=True)
        if not should_auto_check_update():
            if not (self.info and self.info.get("available")):
                self._set_state_text("自动检查已完成，点击可手动检查", "idle")
            return
        if self.checking:
            return
        self.checking = True
        self.button.config(text=self.checking_text, state=tk.DISABLED)
        configure_button_icon(self.root, self.button, "refresh")
        self._set_state_text("正在后台检查更新...", "checking")
        self._set_progress_visible(True, mode="indeterminate", text="检查更新中...")

        def worker():
            info = check_for_update()
            remember_update_check(info)
            return info

        run_loading_task(
            self.root,
            "update_check_worker",
            "正在检查更新...",
            worker,
            lambda info: self._set_check_result(info, silent=True),
            lambda err: self._set_check_error(err, silent=True),
        )

    def on_click(self):
        if self.installing:
            return
        if self.info and self.info.get("available"):
            if self.info.get("asset"):
                self._confirm_and_install(self.info)
            else:
                self._manual_check()
            return
        self._manual_check()

    def _manual_check(self):
        if self.checking:
            return
        self.checking = True
        self.button.config(text=self.checking_text, state=tk.DISABLED)
        configure_button_icon(self.root, self.button, "refresh")
        self._set_state_text("正在连接 GitHub Release...", "checking")
        self._set_progress_visible(True, mode="indeterminate", text="检查更新中...")
        gui_log("正在检查更新...", self.status_var)

        def worker():
            info = check_for_update()
            remember_update_check(info)
            return info

        run_loading_task(
            self.root,
            "update_check_worker",
            "正在检查更新...",
            worker,
            lambda info: self._set_check_result(info, silent=False),
            lambda err: self._set_check_error(err, silent=False),
        )

    def _apply_cached_info(self):
        info = cached_update_info()
        if info and info.get("available"):
            self.info = info
            text = self.available_text or ("有新版本 v%s" % info.get("latest_version"))
            self.button.config(text=text)
            configure_button_icon(self.root, self.button, "update_available")
            self._set_state_text("可更新到 v%s" % info.get("latest_version"), "ready")

    def _set_check_result(self, info, silent, from_cache=False):
        self.checking = False
        self.info = info
        self.button.config(state=tk.NORMAL)
        if not from_cache:
            self._set_progress_visible(False)
        if info.get("available"):
            text = self.available_text or ("有新版本 v%s" % info.get("latest_version"))
            self.button.config(text=text)
            configure_button_icon(self.root, self.button, "update_available")
            self._set_state_text("发现新版本 v%s" % info.get("latest_version"), "ready")
            if not from_cache:
                gui_log("发现新版本 v%s，请点击按钮更新。" % info.get("latest_version"), self.status_var)
            if not silent:
                self._confirm_and_install(info)
            return
        self.button.config(text=self.default_text)
        configure_button_icon(self.root, self.button, "update")
        self._set_state_text("当前已是最新版本 v%s" % info.get("current_version"), "ready")
        if not silent:
            if info.get("missing_asset"):
                messagebox.showwarning("检查更新", "最新 Release 未找到 ExcelMergeFork.exe。")
            else:
                messagebox.showinfo("检查更新", "当前已是最新版本 v%s。" % info.get("current_version"))

    def _set_check_error(self, err, silent):
        self.checking = False
        self.button.config(text=self.default_text, state=tk.NORMAL)
        configure_button_icon(self.root, self.button, "update")
        self._set_state_text("更新检查失败，点击重试", "error")
        self._set_progress_visible(False)
        if not silent:
            messagebox.showerror("检查更新失败", str(err))
        else:
            log("后台检查更新失败: %s" % err, is_error=True)

    def _confirm_and_install(self, info):
        target = get_current_executable()
        if not target:
            messagebox.showinfo("更新提示", "当前是 Python 脚本运行模式，只有打包后的 exe 支持原地更新。")
            return
        msg = (
            "发现新版本 v%s，当前版本 v%s。\n\n"
            "点击确定后会下载新版 exe；下载完成后需要关闭当前窗口，工具会在退出后替换文件。\n"
            "如果正在合并冲突，建议先完成或取消当前合并后再更新。"
        ) % (info.get("latest_version"), info.get("current_version"))
        if not messagebox.askokcancel("发现新版本", msg):
            return
        self.installing = True
        self.button.config(text=self.downloading_text, state=tk.DISABLED)
        configure_button_icon(self.root, self.button, "download")
        self._set_state_text("正在下载 v%s..." % info.get("latest_version"), "checking")
        self._set_progress_visible(True, mode="indeterminate", text="准备下载...")
        gui_log("正在下载 v%s..." % info.get("latest_version"), self.status_var)

        def worker():
            def on_progress(downloaded, total):
                self.root.after(0, lambda d=downloaded, t=total: self._update_download_progress(d, t))

            new_exe = download_update(info, progress_callback=on_progress)
            return make_update_script(new_exe, target, restart_after=False)

        run_loading_task(
            self.root,
            "update_install_worker",
            "正在下载更新...",
            worker,
            self._finish_install,
            self._install_error,
        )

    def _finish_install(self, script):
        self.installing = False
        self._set_progress_visible(True, mode="determinate", value=100, text="下载完成，等待替换...")
        self.button.config(text=(self.available_text or ("有新版本 v%s" % self.info.get("latest_version"))), state=tk.NORMAL)
        configure_button_icon(self.root, self.button, "check", color=UI["success"])
        self._set_state_text("下载完成，准备替换程序", "ready")
        messagebox.showinfo("更新已准备好", "更新包已下载。点击确定后会关闭当前窗口，并自动替换 ExcelMergeFork.exe。")
        launch_update_script(script)
        if self.on_quit:
            self.on_quit()
        else:
            self.root.quit()
            self.root.destroy()

    def _install_error(self, err):
        self.installing = False
        self._set_progress_visible(False)
        self.button.config(text=(self.available_text or ("有新版本 v%s" % self.info.get("latest_version"))), state=tk.NORMAL)
        configure_button_icon(self.root, self.button, "update_available")
        self._set_state_text("更新下载失败，点击可重试", "error")
        if isinstance(err, UpdateError):
            msg = str(err)
        else:
            msg = "更新失败: %s" % err
        gui_log(msg, self.status_var, is_error=True)
        messagebox.showerror("更新失败", msg)

    def _set_progress_visible(self, visible, mode="determinate", value=0, text=""):
        if self.progress_bar is None:
            return
        if visible:
            try:
                self.progress_bar.config(mode=mode)
                if mode == "indeterminate":
                    self.progress_bar.start(12)
                else:
                    self.progress_bar.stop()
                    if self.progress_var is not None:
                        self.progress_var.set(value)
                if self.progress_label is not None:
                    self.progress_label.config(text=text)
                if self.progress_container is not None:
                    self.progress_container.pack(fill=tk.X, pady=(6, 0))
                self.progress_bar.grid()
                if self.progress_label is not None:
                    self.progress_label.grid()
            except Exception:
                pass
        else:
            try:
                self.progress_bar.stop()
                if self.progress_var is not None:
                    self.progress_var.set(0)
                if self.progress_label is not None:
                    self.progress_label.config(text="")
                    self.progress_label.grid_remove()
                self.progress_bar.grid_remove()
                if self.progress_container is not None:
                    self.progress_container.pack_forget()
            except Exception:
                pass

    def _update_download_progress(self, downloaded, total):
        if self.progress_bar is None:
            return
        if total > 0:
            percent = int(downloaded * 100 / total)
            self._set_progress_visible(True, mode="determinate", value=percent, text="下载中 %d%%" % percent)
            self._set_state_text("正在下载更新 %d%%" % percent, "checking")
        else:
            self._set_progress_visible(True, mode="indeterminate", text="下载中...")
            self._set_state_text("正在下载更新...", "checking")

    def _set_state_text(self, text, tone="idle"):
        if self.state_var is not None:
            try:
                self.state_var.set(text)
            except Exception:
                pass
        if self.state_label is not None:
            style_name = {
                "idle": "UpdateIdle.TLabel",
                "checking": "UpdateChecking.TLabel",
                "ready": "UpdateReady.TLabel",
                "error": "UpdateError.TLabel",
            }.get(tone, "UpdateIdle.TLabel")
            try:
                self.state_label.configure(style=style_name)
            except Exception:
                pass
        if self.state_icon is not None:
            icon_name = {
                "idle": "update",
                "checking": "update_checking",
                "ready": "update_available",
                "error": "update_error",
            }.get(tone, "update")
            try:
                img = get_ui_icon(self.root, icon_name, size=16)
                self.state_icon.configure(image=img)
                self.state_icon._excel_merge_icon = img
            except Exception:
                pass
