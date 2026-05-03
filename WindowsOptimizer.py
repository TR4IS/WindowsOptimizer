"""
Windows Performance & Gaming Optimizer
Refined Dark Spring Orange Edition
"""

import customtkinter as ctk
from tkinter import messagebox
import threading
import subprocess
import os
import sys
import ctypes
import winreg
import time
import glob
import webbrowser
import hashlib
import json
import urllib.request
import tempfile

# ─── App Info ────────────────────────────────────────────────────────────────
VERSION    = "1.3.4"
GITHUB_URL = "https://github.com/TR4IS/WindowsOptimizer"
UPDATE_URL = "https://raw.githubusercontent.com/TR4IS/WindowsOptimizer/main/docs/version.json"

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def is_file_ready(file_path):
    """Check if a file is completely written and unlocked."""
    try:
        if not os.path.exists(file_path): return False
        with open(file_path, 'a'): pass
        return True
    except (PermissionError, OSError): return False

# ─── Appearance ─────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ─── Dark Spring Orange Palette ──────────────────────────────────────────────
# Backgrounds: Deep, dark orange-tinted black for a premium look
BASE_BG      = "#0a0500"  # Darkest Burnt Orange
SIDEBAR_BG   = "#140c00"  # Deep Clay
PANEL_BG     = "#1f1400"  # Warm Charcoal Orange
BORDER       = "#2e1d00"  # Subtle separator

# Accents: Vibrant Spring Orange
ACCENT       = "#ff8c00"  # Dark Orange / Spring Orange
ACCENT_HOVER = "#e67e00"
ACCENT_SOFT  = "#4d2b00"  # Muted glow

# Text: Warm Grays with cream undertones
TEXT_MAIN    = "#fff8f0"  # Cream White
TEXT_SUB     = "#d1c7b8"  # Warm Gray
TEXT_MUTED   = "#665c52"  # Muted Earth

# Feedback
SUCCESS      = "#ffa500"  # Orange
WARNING      = "#ffcc00"  # Warm Gold
DANGER       = "#ff4d00"  # Bright Orange-Red

# ─── Admin check ─────────────────────────────────────────────────────────────
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False

def run_cmd(cmd, shell=False, timeout=30):
    try:
        result = subprocess.run(
            cmd, shell=shell, capture_output=True, text=True, timeout=timeout
        )
        return result.returncode == 0, result.stdout + result.stderr
    except Exception as e:
        return False, str(e)

# ─── Optimization functions ───────────────────────────────────────────────────

def clean_temp_files(log):
    log("Cleaning temporary files...")
    paths = [
        os.environ.get("TEMP", ""),
        os.environ.get("TMP", ""),
        os.path.join(os.environ.get("SystemRoot", "C:\\Windows"), "Temp"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Temp"),
    ]
    total = 0
    for p in paths:
        if not p or not os.path.isdir(p):
            continue
        for root, dirs, files in os.walk(p):
            for f in files:
                fp = os.path.join(root, f)
                try:
                    size = os.path.getsize(fp)
                    os.remove(fp)
                    total += size 
                except Exception:
                    pass
    log(f"   Removed {total // (1024*1024)} MB of junk data")


def clean_prefetch(log):
    log("Cleaning Prefetch...")
    prefetch_dir = os.path.join(os.environ.get("SystemRoot", "C:\\Windows"), "Prefetch")
    removed = 0
    if os.path.isdir(prefetch_dir):
        for fp in glob.glob(os.path.join(prefetch_dir, "*")):
            try:
                os.remove(fp)
                removed += 1
            except Exception:
                pass
    log(f"   Cleared {removed} prefetch entries")


def empty_recycle_bin(log):
    log("Emptying Recycle Bin...")
    try:
        ctypes.windll.shell32.SHEmptyRecycleBinW(None, None, 0x0007)
        log("   Recycle Bin emptied")
    except Exception as e:
        log(f"   Error: {e}")


def set_high_performance_power(log):
    log("Setting High Performance power plan...")
    ok, _ = run_cmd(["powercfg", "/setactive", "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"])
    if not ok:
        run_cmd(["powercfg", "/duplicatescheme", "e9a42b02-d5df-448d-aa00-03f14749eb61"])
        ok2, _ = run_cmd(["powercfg", "/setactive", "e9a42b02-d5df-448d-aa00-03f14749eb61"])
        log("   Ultimate Performance active" if ok2 else "   High Performance active")
    else:
        log("   High Performance active")


def disable_startup_programs(log):
    log("Disabling common startup bloat...")
    bloat = [
        "OneDrive", "Skype", "Spotify", "Discord", "Teams",
        "AdobeUpdater", "GoogleUpdate", "iTunesHelper",
        "SteamService", "EpicGamesLauncher",
    ]
    key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    disabled = []
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS) as key:
            for name in bloat:
                try:
                    winreg.DeleteValue(key, name)
                    disabled.append(name)
                except Exception:
                    pass
    except Exception:
        pass
        
    if disabled:
        log(f"   Disabled: {', '.join(disabled)}")
    else:
        log("   No startup bloat found")


def disable_visual_effects(log):
    log("Disabling unnecessary visual effects...")
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects",
            0, winreg.KEY_SET_VALUE
        ) as key:
            winreg.SetValueEx(key, "VisualFXSetting", 0, winreg.REG_DWORD, 2)
        log("   Visual effects optimized for performance")
    except Exception as e:
        log(f"   Error: {e}")


def set_game_mode(log):
    log("Enabling Windows Game Mode...")
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\GameBar") as key:
            winreg.SetValueEx(key, "AutoGameModeEnabled", 0, winreg.REG_DWORD, 1)
            winreg.SetValueEx(key, "AllowAutoGameMode", 0, winreg.REG_DWORD, 1)
        log("   Game Mode enabled")
    except Exception as e:
        log(f"   Error: {e}")


def disable_xbox_game_bar(log):
    log("Disabling Xbox DVR and Overlay...")
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\GameDVR") as key:
            winreg.SetValueEx(key, "AppCaptureEnabled", 0, winreg.REG_DWORD, 0)
            
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"System\GameConfigStore") as key2:
            winreg.SetValueEx(key2, "GameDVR_Enabled", 0, winreg.REG_DWORD, 0)
        log("   Xbox services disabled")
    except Exception as e:
        log(f"   Error: {e}")


def set_gpu_scheduling(log):
    log("Enabling Hardware GPU Scheduling...")
    try:
        with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\GraphicsDrivers") as key:
            winreg.SetValueEx(key, "HwSchMode", 0, winreg.REG_DWORD, 2)
        log("   HAGS enabled (restart required)")
    except Exception as e:
        log(f"   Error: {e}")


def flush_dns(log):
    log("Flushing DNS cache...")
    ok, _ = run_cmd(["ipconfig", "/flushdns"])
    log("   DNS flushed" if ok else "   Failed to flush DNS")


def disable_search_indexing(log):
    log("Stopping Search indexing service...")
    run_cmd(["sc", "stop", "WSearch"])
    ok, _ = run_cmd(["sc", "config", "WSearch", "start=", "disabled"])
    log("   Indexing disabled" if ok else "   Could not disable service")


def clean_winsxs(log):
    log("Running DISM component cleanup...")
    log("   This process takes time. Please wait...")
    ok, out = run_cmd(
        ["dism", "/online", "/cleanup-image", "/startcomponentcleanup", "/resetbase"],
        timeout=None
    )
    log("   DISM cleanup successful" if ok else f"   Error: {out[:120].strip()}")


def adjust_timer_resolution(log):
    log("Optimizing CPU scheduling...")
    try:
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SYSTEM\CurrentControlSet\Control\PriorityControl",
            0, winreg.KEY_SET_VALUE
        ) as key:
            winreg.SetValueEx(key, "Win32PrioritySeparation", 0, winreg.REG_DWORD, 38)
        log("   Priority tuned for active applications")
    except Exception as e:
        log(f"   Error: {e}")


def disable_telemetry(log):
    log("Reducing telemetry and data collection...")
    cmds = [
        ["sc", "stop", "DiagTrack"],
        ["sc", "config", "DiagTrack", "start=", "disabled"],
        ["sc", "stop", "dmwappushservice"],
        ["sc", "config", "dmwappushservice", "start=", "disabled"],
    ]
    for cmd in cmds:
        run_cmd(cmd)
    try:
        with winreg.CreateKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Policies\Microsoft\Windows\DataCollection"
        ) as key:
            winreg.SetValueEx(key, "AllowTelemetry", 0, winreg.REG_DWORD, 0)
    except Exception:
        pass
    log("   Telemetry disabled")


def optimize_network(log):
    log("Calibrating network for low latency...")
    cmds = [
        ["netsh", "int", "tcp", "set", "global", "autotuninglevel=normal"],
        ["netsh", "int", "tcp", "set", "global", "chimney=disabled"],
        ["netsh", "int", "tcp", "set", "global", "dca=enabled"],
        ["netsh", "int", "tcp", "set", "global", "netdma=enabled"],
        ["netsh", "int", "tcp", "set", "global", "ecncapability=disabled"],
        ["netsh", "int", "tcp", "set", "global", "timestamps=disabled"],
        ["netsh", "int", "tcp", "set", "supplemental", "template=internet", "congestionprovider=ctcp"],
    ]
    for cmd in cmds:
        run_cmd(cmd)
    log("   Network stack calibrated")


def run_disk_cleanup(log):
    log("Starting automated Disk Cleanup...")
    run_cmd(["cleanmgr", "/sagerun:1"])
    log("   Cleanup utility launched")


# ─── Revert functions ──────────────────────────────────────────────────────────

def revert_power_plan(log):
    log("Restoring Balanced power plan...")
    ok, _ = run_cmd(["powercfg", "/setactive", "381b4222-f694-41f0-9685-ff5bb260df2e"])
    log("   Balanced plan set" if ok else "   Failed to restore plan")

def revert_timer_resolution(log):
    log("Restoring default CPU scheduling...")
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\PriorityControl", 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, "Win32PrioritySeparation", 0, winreg.REG_DWORD, 2)
        log("   Scheduling restored to default")
    except Exception as e: log(f"   Error: {e}")

def revert_game_mode(log):
    log("Disabling Game Mode...")
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\GameBar") as key:
            winreg.SetValueEx(key, "AutoGameModeEnabled", 0, winreg.REG_DWORD, 0)
            winreg.SetValueEx(key, "AllowAutoGameMode", 0, winreg.REG_DWORD, 0)
        log("   Game Mode disabled")
    except Exception as e: log(f"   Error: {e}")

def revert_xbox_game_bar(log):
    log("Restoring Xbox services...")
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\GameDVR") as key:
            winreg.SetValueEx(key, "AppCaptureEnabled", 0, winreg.REG_DWORD, 1)
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"System\GameConfigStore") as key2:
            winreg.SetValueEx(key2, "GameDVR_Enabled", 0, winreg.REG_DWORD, 1)
        log("   Xbox services enabled")
    except Exception as e: log(f"   Error: {e}")

def revert_gpu_scheduling(log):
    log("Disabling GPU Scheduling...")
    try:
        with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\GraphicsDrivers") as key:
            winreg.SetValueEx(key, "HwSchMode", 0, winreg.REG_DWORD, 1)
        log("   HAGS disabled")
    except Exception as e: log(f"   Error: {e}")

def revert_visual_effects(log):
    log("Restoring default visual effects...")
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects", 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, "VisualFXSetting", 0, winreg.REG_DWORD, 0)
        log("   Visual effects restored")
    except Exception as e: log(f"   Error: {e}")

def revert_search_indexing(log):
    log("Restoring Search indexing...")
    run_cmd(["sc", "config", "WSearch", "start=", "delayed-auto"])
    ok, _ = run_cmd(["sc", "start", "WSearch"])
    log("   Indexing enabled" if ok else "   Could not restore service")

def revert_telemetry(log):
    log("Restoring telemetry services...")
    run_cmd(["sc", "config", "DiagTrack", "start=", "auto"])
    run_cmd(["sc", "start", "DiagTrack"])
    try:
        with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Policies\Microsoft\Windows\DataCollection") as key:
            winreg.DeleteValue(key, "AllowTelemetry")
    except Exception: pass
    log("   Telemetry restored")

def revert_network_optimization(log):
    log("Restoring default network settings...")
    cmds = [
        ["netsh", "int", "tcp", "set", "global", "autotuninglevel=normal"],
        ["netsh", "int", "tcp", "set", "global", "rss=enabled"],
        ["netsh", "int", "tcp", "set", "global", "timestamps=disabled"],
    ]
    for cmd in cmds: run_cmd(cmd)
    log("   Network restored")

# ─── Task Data ───────────────────────────────────────────────────────────────

TASKS = {
    "System Maintenance": {
        "funcs": [clean_temp_files, clean_prefetch, empty_recycle_bin],
        "desc": "Removes temporary files, prefetch data, and empties the recycle bin to free up disk space."
    },
    "Performance Tuning": {
        "funcs": [set_high_performance_power, adjust_timer_resolution],
        "desc": "Activates the High Performance power plan and tunes CPU scheduling for active applications."
    },
    "Gaming Enhancements": {
        "funcs": [set_game_mode, disable_xbox_game_bar, set_gpu_scheduling, disable_visual_effects],
        "desc": "Enables Game Mode, HAGS, and disables Xbox DVR and heavy visual effects for maximum FPS."
    },
    "Service Optimization": {
        "funcs": [disable_startup_programs, disable_search_indexing, disable_telemetry],
        "desc": "Disables startup bloat, Search indexing, and telemetry services to reduce background CPU usage."
    },
    "Network Calibration": {
        "funcs": [flush_dns, optimize_network],
        "desc": "Optimizes the TCP stack and flushes DNS to reduce latency and jitter in online games."
    },
    "Advanced Cleanup": {
        "funcs": [clean_winsxs, run_disk_cleanup],
        "desc": "Performs deep DISM component cleanup and launches the system disk cleanup utility."
    },
}

REVERT_TASKS = {
    "Performance Tuning": [revert_power_plan, revert_timer_resolution],
    "Gaming Enhancements": [revert_game_mode, revert_xbox_game_bar, revert_gpu_scheduling, revert_visual_effects],
    "Service Optimization": [revert_search_indexing, revert_telemetry],
    "Network Calibration": [revert_network_optimization],
}

# ─── GUI Components ──────────────────────────────────────────────────────────

class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tip_window or not self.text:
            return
        x = self.widget.winfo_rootx() + 25
        y = self.widget.winfo_rooty() + 40
        self.tip_window = tw = ctk.CTkToplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry(f"+{x}+{y}")
        tw.attributes("-topmost", True)
        
        label = ctk.CTkLabel(
            tw, text=self.text, justify="left",
            bg_color=PANEL_BG, text_color=TEXT_SUB,
            padx=10, pady=5, font=ctk.CTkFont(size=12),
            corner_radius=4
        )
        label.pack()

    def hide_tip(self, event=None):
        tw = self.tip_window
        self.tip_window = None
        if tw:
            tw.destroy()

class OptimizerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"Windows Optimizer v{VERSION}")
        self.geometry("1100x800")
        self.configure(fg_color=BASE_BG)
        self.running = False
        self.checks = {}
        
        # Set icon
        try:
            icon_path = resource_path('WindowsOptimizer.ico')
            if os.path.exists(icon_path):
                self.after(201, lambda: self.iconbitmap(icon_path))
        except Exception:
            pass

        self._build_ui()
        self.check_for_updates()

    def _build_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ── Sidebar ──
        self.sidebar = ctk.CTkFrame(self, width=340, corner_radius=0, fg_color=SIDEBAR_BG, border_width=1, border_color=BORDER)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(4, weight=1)

        ctk.CTkLabel(
            self.sidebar, text="Optimizer",
            font=ctk.CTkFont(size=30, weight="normal"), text_color=ACCENT
        ).grid(row=0, column=0, padx=20, pady=(40, 5))

        admin_color = SUCCESS if is_admin() else DANGER
        admin_text  = "Admin Mode Active" if is_admin() else "Limited Mode"
        ctk.CTkLabel(
            self.sidebar, text=admin_text,
            font=ctk.CTkFont(size=13, weight="normal"), text_color=admin_color
        ).grid(row=1, column=0, padx=20, pady=(0, 40))

        # Checkboxes
        self.scrollable_frame = ctk.CTkScrollableFrame(
            self.sidebar, label_text="Tweaks", 
            fg_color="transparent", label_text_color=ACCENT,
            border_width=0, scrollbar_button_color=ACCENT_SOFT
        )
        self.scrollable_frame.grid(row=2, column=0, padx=15, pady=0, sticky="nsew")
        self.scrollable_frame.grid_columnconfigure(0, weight=1)

        for group, data in TASKS.items():
            var = ctk.BooleanVar(value=True)
            self.checks[group] = var
            cb = ctk.CTkCheckBox(
                self.scrollable_frame, text=group, variable=var,
                font=ctk.CTkFont(size=15, weight="normal"), 
                border_width=1, corner_radius=4,
                fg_color=ACCENT, hover_color=ACCENT_HOVER,
                text_color=TEXT_SUB, text_color_disabled=TEXT_MUTED
            )
            cb.pack(fill="x", padx=15, pady=12)
            ToolTip(cb, data["desc"])

        # Selection Buttons
        self.sel_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.sel_frame.grid(row=3, column=0, padx=20, pady=30)

        self.all_btn = ctk.CTkButton(
            self.sel_frame, text="All", height=32, width=100,
            fg_color=PANEL_BG, border_width=1, border_color=BORDER,
            hover_color=BORDER, font=ctk.CTkFont(size=12, weight="normal"),
            text_color=TEXT_SUB, command=self._select_all
        )
        self.all_btn.pack(side="left", padx=8)

        self.none_btn = ctk.CTkButton(
            self.sel_frame, text="None", height=32, width=100,
            fg_color=PANEL_BG, border_width=1, border_color=BORDER,
            hover_color=BORDER, font=ctk.CTkFont(size=12, weight="normal"),
            text_color=TEXT_SUB, command=self._select_none
        )
        self.none_btn.pack(side="left", padx=8)

        # Sidebar Footer
        self.footer = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.footer.grid(row=5, column=0, padx=20, pady=30, sticky="s")

        self.gh_btn = ctk.CTkButton(
            self.footer, text="Open GitHub", height=36,
            fg_color="transparent", border_width=1, border_color=BORDER,
            hover_color=PANEL_BG, font=ctk.CTkFont(size=13, weight="normal"),
            text_color=TEXT_SUB, command=self._open_github
        )
        self.gh_btn.pack(fill="x", pady=6)

        self.upd_btn = ctk.CTkButton(
            self.footer, text="Check Updates", height=36,
            fg_color="transparent", border_width=1, border_color=ACCENT,
            text_color=ACCENT, hover_color=PANEL_BG,
            font=ctk.CTkFont(size=13, weight="normal"),
            command=lambda: self.check_for_updates(manual=True)
        )
        self.upd_btn.pack(fill="x", pady=6)

        # ── Main Area ──
        self.main_frame = ctk.CTkFrame(self, fg_color=BASE_BG)
        self.main_frame.grid(row=0, column=1, padx=40, pady=40, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)

        self.header_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 25))
        
        ctk.CTkLabel(
            self.header_frame, text="Operation Console",
            font=ctk.CTkFont(size=22, weight="normal"), text_color=TEXT_MAIN
        ).pack(side="left")

        self.clear_btn = ctk.CTkButton(
            self.header_frame, text="Clear Console", width=120, height=30,
            fg_color=PANEL_BG, border_width=1, border_color=BORDER,
            hover_color=BORDER, font=ctk.CTkFont(size=12, weight="normal"),
            text_color=TEXT_SUB, command=self._clear_log
        )
        self.clear_btn.pack(side="right")

        self.log_box = ctk.CTkTextbox(
            self.main_frame, 
            font=ctk.CTkFont(family="Consolas", size=13),
            fg_color=SIDEBAR_BG, border_width=1, border_color=BORDER,
            text_color=TEXT_MAIN, corner_radius=4
        )
        self.log_box.grid(row=1, column=0, sticky="nsew")
        
        self.log_box._textbox.tag_config("ok",      foreground=SUCCESS)
        self.log_box._textbox.tag_config("err",     foreground=DANGER)
        self.log_box._textbox.tag_config("section", foreground=ACCENT)
        self.log_box._textbox.tag_config("info",    foreground=WARNING)

        self.actions = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.actions.grid(row=2, column=0, sticky="ew", pady=(30, 0))

        self.run_btn = ctk.CTkButton(
            self.actions, text="Start Optimization", height=60,
            font=ctk.CTkFont(size=17, weight="normal"),
            fg_color=ACCENT, hover_color=ACCENT_HOVER, text_color=BASE_BG,
            command=self._run, corner_radius=4
        )
        self.run_btn.pack(side="left", fill="x", expand=True, padx=(0, 15))

        self.undo_btn = ctk.CTkButton(
            self.actions, text="Restore Defaults", height=60,
            font=ctk.CTkFont(size=17, weight="normal"),
            fg_color=PANEL_BG, border_width=1, border_color=BORDER,
            hover_color=BORDER, text_color=WARNING,
            command=self._undo, corner_radius=4
        )
        self.undo_btn.pack(side="left", fill="x", expand=True, padx=(15, 0))

        self.progress = ctk.CTkProgressBar(
            self.main_frame, mode="indeterminate", 
            fg_color=PANEL_BG, progress_color=ACCENT, height=2
        )
        self.progress.grid(row=3, column=0, sticky="ew", pady=(25, 0))
        self.progress.set(0)

        self._log("System ready. Hover over tweaks for details.", "info")
        if not is_admin():
            self._log("Warning: Admin rights recommended for full effect.", "err")

    # ── helpers ──────────────────────────────────────────────────────────────

    def check_for_updates(self, manual=False):
        def _check():
            try:
                req = urllib.request.Request(UPDATE_URL, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=10) as response:
                    if response.status != 200:
                        raise Exception(f"HTTP {response.status}")
                    data = json.loads(response.read().decode())
                
                remote_version = data.get("version")
                download_url = data.get("url")
                expected_hash = data.get("sha256")

                if remote_version and remote_version > VERSION:
                    if messagebox.askyesno("Update", f"Version {remote_version} is available. Install now?"):
                        self._log(f"Downloading update {remote_version}...", "info")
                        
                        ext = os.path.splitext(download_url)[1] or (".exe" if getattr(sys, 'frozen', False) else ".py")
                        update_path = os.path.join(tempfile.gettempdir(), f"WindowsOptimizer_update{ext}")
                        
                        req_dl = urllib.request.Request(download_url, headers={'User-Agent': 'Mozilla/5.0'})
                        with urllib.request.urlopen(req_dl, timeout=30) as response:
                            sha256_hash = hashlib.sha256()
                            with open(update_path, 'wb') as f:
                                while True:
                                    chunk = response.read(8192)
                                    if not chunk: break
                                    f.write(chunk)
                                    sha256_hash.update(chunk)
                        
                        if not is_file_ready(update_path):
                            self._log("Error: Download failed.", "err")
                            return

                        if expected_hash and sha256_hash.hexdigest().lower() != expected_hash.lower():
                            self._log("Security: Hash mismatch.", "err")
                            if os.path.exists(update_path): os.remove(update_path)
                            messagebox.showerror("Error", "Update verification failed.")
                            return
                        
                        self._log("Starting update...", "ok")
                        os.startfile(update_path)
                        self.after(0, self.destroy)
                        os._exit(0)
                elif manual:
                    messagebox.showinfo("Update", f"You are running version {VERSION}.")
            except Exception as e:
                self._log(f"Check failed: {e}", "err")

        threading.Thread(target=_check, daemon=True).start()

    def _undo(self):
        if self.running: return
        selected = [g for g, v in self.checks.items() if v.get() and g in REVERT_TASKS]
        if not selected:
            messagebox.showwarning("Notice", "Select categories to restore.")
            return
        if not messagebox.askyesno("Confirm", "Restore defaults for selected categories?"):
            return
        self.running = True
        self.undo_btn.configure(state="disabled", text="Restoring...")
        self.progress.start()
        threading.Thread(target=self._undo_worker, args=(selected,), daemon=True).start()

    def _undo_worker(self, groups):
        t0 = time.time()
        self._log("\n" + "-"*40, "section")
        self._log("RESTORING DEFAULTS", "section")
        self._log("-"*40 + "\n", "section")

        for group in groups:
            self._log(f"Restoring: {group}", "section")
            for fn in REVERT_TASKS[group]:
                try:
                    fn(self._log)
                except Exception as e:
                    self._log(f"   Error: {e}", "err")

        elapsed = time.time() - t0
        self._log("\n" + "-"*40, "section")
        self._log(f"RESTORE COMPLETE ({elapsed:.1f}s)", "ok")
        self._log("-"*40, "section")

        def finalize():
            self.running = False
            self.progress.stop()
            self.undo_btn.configure(state="normal", text="Restore Defaults")
        self.after(0, finalize)

    def _open_github(self):
        webbrowser.open(GITHUB_URL)

    def _log(self, msg, tag=""):
        def update_ui():
            self.log_box.configure(state="normal")
            self.log_box.insert("end", msg + "\n", tag)
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self.after(0, update_ui)

    def _clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    def _select_all(self):
        for v in self.checks.values():
            v.set(True)

    def _select_none(self):
        for v in self.checks.values():
            v.set(False)

    def _run(self):
        if self.running:
            return
        selected = [g for g, v in self.checks.items() if v.get()]
        if not selected:
            messagebox.showwarning("Notice", "Select at least one category.")
            return
        self.running = True
        self.run_btn.configure(state="disabled", text="Processing...")
        self.progress.start()
        threading.Thread(target=self._worker, args=(selected,), daemon=True).start()

    def _worker(self, groups):
        t0 = time.time()
        self._log("\n" + "-"*40, "section")
        self._log("OPTIMIZATION BEGUN", "section")
        self._log("-"*40 + "\n", "section")

        for group in groups:
            self._log(f"Processing: {group}", "section")
            for fn in TASKS[group]["funcs"]:
                try:
                    fn(self._log)
                except Exception as e:
                    self._log(f"   Error: {e}", "err")

        elapsed = time.time() - t0
        self._log("\n" + "-"*40, "section")
        self._log(f"PROCESS COMPLETE ({elapsed:.1f}s)", "ok")
        self._log("Restart your device for all changes to apply.", "info")
        self._log("-"*40, "section")

        def finalize():
            self.running = False
            self.progress.stop()
            self.run_btn.configure(state="normal", text="Start Optimization")
        self.after(0, finalize)

# ─── Entry point ─────────────────────────────────────────────────────────────

mutex_handle = None

def check_lock():
    global mutex_handle
    try:
        mutex_name = "TR4IS_Windows_Optimizer_Instance_Lock"
        mutex_handle = ctypes.windll.kernel32.CreateMutexW(None, False, mutex_name)
        if ctypes.windll.kernel32.GetLastError() == 183: # ERROR_ALREADY_EXISTS
            return False
        return True
    except Exception:
        return True

if __name__ == "__main__":
    if sys.platform != "win32":
        sys.exit(1)

    if not check_lock():
        time.sleep(0.5)
        if not check_lock():
            ctypes.windll.user32.MessageBoxW(0, "Windows Optimizer is already running.", "Instance Active", 0x40 | 0x0)
            sys.exit(0)

    if not is_admin():
        try:
            args_string = " ".join(f'"{arg}"' for arg in sys.argv)
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, args_string, None, 1)
        except Exception:
            pass
        sys.exit(0)

    app = OptimizerApp()
    app.mainloop()
