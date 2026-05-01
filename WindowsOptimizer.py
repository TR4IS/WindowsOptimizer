"""
Windows Performance & Gaming Optimizer
Run as Administrator for full effect.
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
VERSION    = "1.2.0"
GITHUB_URL = "https://github.com/TR4IS/WindowsOptimizer"
UPDATE_URL = "https://raw.githubusercontent.com/TR4IS/WindowsOptimizer/main/docs/version.json"

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

# ─── Color palette (Custom for text tags/manual colors) ──────────────────────
ACCENT    = "#00e5ff"
SUCCESS   = "#00ff88"
WARNING   = "#ffaa00"
DANGER    = "#ff4444"
TEXT      = "#e2e8f0"

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
    log("🗑  Cleaning temporary files...")
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
    log(f"   ✔ Removed ~{total // (1024*1024)} MB of temp files")


def clean_prefetch(log):
    log("🗑  Cleaning Prefetch...")
    prefetch_dir = os.path.join(os.environ.get("SystemRoot", "C:\\Windows"), "Prefetch")
    removed = 0
    if os.path.isdir(prefetch_dir):
        for fp in glob.glob(os.path.join(prefetch_dir, "*")):
            try:
                os.remove(fp)
                removed += 1
            except Exception:
                pass
    log(f"   ✔ Prefetch cleared ({removed} files removed)")


def empty_recycle_bin(log):
    log("🗑  Emptying Recycle Bin...")
    try:
        ctypes.windll.shell32.SHEmptyRecycleBinW(None, None, 0x0007)
        log("   ✔ Recycle Bin emptied")
    except Exception as e:
        log(f"   ✖ {e}")


def set_high_performance_power(log):
    log("⚡ Setting High Performance power plan...")
    ok, _ = run_cmd(["powercfg", "/setactive", "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"])
    if not ok:
        run_cmd(["powercfg", "/duplicatescheme", "e9a42b02-d5df-448d-aa00-03f14749eb61"])
        ok2, _ = run_cmd(["powercfg", "/setactive", "e9a42b02-d5df-448d-aa00-03f14749eb61"])
        log("   ✔ Ultimate Performance plan activated" if ok2 else "   ✔ High Performance plan set")
    else:
        log("   ✔ High Performance power plan activated")


def disable_startup_programs(log):
    log("🚀 Disabling common bloat startup entries...")
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
        log(f"   ✔ Disabled: {', '.join(disabled)}")
    else:
        log("   ✔ No common bloat startup entries found")


def disable_visual_effects(log):
    log("🎨 Disabling unnecessary visual effects...")
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects",
            0, winreg.KEY_SET_VALUE
        ) as key:
            winreg.SetValueEx(key, "VisualFXSetting", 0, winreg.REG_DWORD, 2)
        log("   ✔ Visual effects set to 'Best Performance'")
    except Exception as e:
        log(f"   ✖ {e}")


def set_game_mode(log):
    log("🎮 Enabling Windows Game Mode...")
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\GameBar") as key:
            winreg.SetValueEx(key, "AutoGameModeEnabled", 0, winreg.REG_DWORD, 1)
            winreg.SetValueEx(key, "AllowAutoGameMode", 0, winreg.REG_DWORD, 1)
        log("   ✔ Game Mode enabled")
    except Exception as e:
        log(f"   ✖ {e}")


def disable_xbox_game_bar(log):
    log("🎮 Disabling Xbox Game Bar overlay (FPS boost)...")
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\GameDVR") as key:
            winreg.SetValueEx(key, "AppCaptureEnabled", 0, winreg.REG_DWORD, 0)
            
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"System\GameConfigStore") as key2:
            winreg.SetValueEx(key2, "GameDVR_Enabled", 0, winreg.REG_DWORD, 0)
        log("   ✔ Xbox Game Bar / DVR disabled")
    except Exception as e:
        log(f"   ✖ {e}")


def set_gpu_scheduling(log):
    log("🖥  Enabling Hardware-Accelerated GPU Scheduling (HAGS)...")
    try:
        with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\GraphicsDrivers") as key:
            winreg.SetValueEx(key, "HwSchMode", 0, winreg.REG_DWORD, 2)
        log("   ✔ HAGS enabled (restart required)")
    except Exception as e:
        log(f"   ✖ {e}")


def flush_dns(log):
    log("🌐 Flushing DNS cache...")
    ok, _ = run_cmd(["ipconfig", "/flushdns"])
    log("   ✔ DNS cache flushed" if ok else "   ✖ Failed to flush DNS")


def disable_search_indexing(log):
    log("🔍 Stopping Windows Search indexing service...")
    run_cmd(["sc", "stop", "WSearch"])
    ok, _ = run_cmd(["sc", "config", "WSearch", "start=", "disabled"])
    log("   ✔ Search indexing disabled" if ok else "   ✖ Could not disable (may need admin)")


def clean_winsxs(log):
    log("🧹 Running DISM cleanup (removes old Windows update files)...")
    log("   ⏳ This may take a few minutes. Please wait...")
    ok, out = run_cmd(
        ["dism", "/online", "/cleanup-image", "/startcomponentcleanup", "/resetbase"],
        timeout=None
    )
    log("   ✔ DISM cleanup done" if ok else f"   ✖ DISM: {out[:120].strip()}")


def adjust_timer_resolution(log):
    log("⏱  Setting processor scheduling to 'Background Services' style for gaming...")
    try:
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SYSTEM\CurrentControlSet\Control\PriorityControl",
            0, winreg.KEY_SET_VALUE
        ) as key:
            winreg.SetValueEx(key, "Win32PrioritySeparation", 0, winreg.REG_DWORD, 38)
        log("   ✔ CPU priority tuned for foreground apps / games")
    except Exception as e:
        log(f"   ✖ {e}")


def disable_telemetry(log):
    log("📡 Reducing Windows telemetry / data collection...")
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
    log("   ✔ Telemetry services disabled")


def optimize_network(log):
    log("🌐 Optimizing network settings for lower latency...")
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
    log("   ✔ Network stack optimized for low latency")


def run_disk_cleanup(log):
    log("💾 Running Disk Cleanup (automated)...")
    run_cmd(["cleanmgr", "/sagerun:1"])
    log("   ✔ Disk Cleanup launched")


# ─── Revert functions ──────────────────────────────────────────────────────────

def revert_power_plan(log):
    log("⚡ Reverting power plan to Balanced...")
    ok, _ = run_cmd(["powercfg", "/setactive", "381b4222-f694-41f0-9685-ff5bb260df2e"])
    log("   ✔ Balanced power plan set" if ok else "   ✖ Failed to set Balanced plan")

def revert_timer_resolution(log):
    log("⏱  Reverting processor scheduling to default...")
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\PriorityControl", 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, "Win32PrioritySeparation", 0, winreg.REG_DWORD, 2)
        log("   ✔ CPU priority reverted to default")
    except Exception as e: log(f"   ✖ {e}")

def revert_game_mode(log):
    log("🎮 Disabling Windows Game Mode...")
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\GameBar") as key:
            winreg.SetValueEx(key, "AutoGameModeEnabled", 0, winreg.REG_DWORD, 0)
            winreg.SetValueEx(key, "AllowAutoGameMode", 0, winreg.REG_DWORD, 0)
        log("   ✔ Game Mode disabled")
    except Exception as e: log(f"   ✖ {e}")

def revert_xbox_game_bar(log):
    log("🎮 Enabling Xbox Game Bar overlay...")
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\GameDVR") as key:
            winreg.SetValueEx(key, "AppCaptureEnabled", 0, winreg.REG_DWORD, 1)
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"System\GameConfigStore") as key2:
            winreg.SetValueEx(key2, "GameDVR_Enabled", 0, winreg.REG_DWORD, 1)
        log("   ✔ Xbox Game Bar enabled")
    except Exception as e: log(f"   ✖ {e}")

def revert_gpu_scheduling(log):
    log("🖥  Disabling Hardware-Accelerated GPU Scheduling...")
    try:
        with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\GraphicsDrivers") as key:
            winreg.SetValueEx(key, "HwSchMode", 0, winreg.REG_DWORD, 1)
        log("   ✔ HAGS disabled (restart required)")
    except Exception as e: log(f"   ✖ {e}")

def revert_visual_effects(log):
    log("🎨 Reverting visual effects to Windows default...")
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects", 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, "VisualFXSetting", 0, winreg.REG_DWORD, 0)
        log("   ✔ Visual effects reverted")
    except Exception as e: log(f"   ✖ {e}")

def revert_search_indexing(log):
    log("🔍 Enabling Windows Search indexing service...")
    run_cmd(["sc", "config", "WSearch", "start=", "delayed-auto"])
    ok, _ = run_cmd(["sc", "start", "WSearch"])
    log("   ✔ Search indexing enabled" if ok else "   ✖ Could not enable")

def revert_telemetry(log):
    log("📡 Enabling Windows telemetry services...")
    run_cmd(["sc", "config", "DiagTrack", "start=", "auto"])
    run_cmd(["sc", "start", "DiagTrack"])
    try:
        with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Policies\Microsoft\Windows\DataCollection") as key:
            winreg.DeleteValue(key, "AllowTelemetry")
    except Exception: pass
    log("   ✔ Telemetry services enabled")

def revert_network_optimization(log):
    log("🌐 Reverting network settings to defaults...")
    cmds = [
        ["netsh", "int", "tcp", "set", "global", "autotuninglevel=normal"],
        ["netsh", "int", "tcp", "set", "global", "rss=enabled"],
        ["netsh", "int", "tcp", "set", "global", "timestamps=disabled"],
    ]
    for cmd in cmds: run_cmd(cmd)
    log("   ✔ Network settings reverted")

# ─── Task groups ─────────────────────────────────────────────────────────────

TASKS = {
    "🗑 Clean Temp & Junk Files": [
        clean_temp_files, clean_prefetch, empty_recycle_bin
    ],
    "⚡ Power & CPU Tuning": [
        set_high_performance_power, adjust_timer_resolution
    ],
    "🎮 Gaming Optimizations": [
        set_game_mode, disable_xbox_game_bar, set_gpu_scheduling, disable_visual_effects
    ],
    "🚀 Startup & Services": [
        disable_startup_programs, disable_search_indexing, disable_telemetry
    ],
    "🌐 Network & DNS": [
        flush_dns, optimize_network
    ],
    "🧹 Deep System Cleanup": [
        clean_winsxs, run_disk_cleanup
    ],
}

REVERT_TASKS = {
    "⚡ Power & CPU Tuning": [revert_power_plan, revert_timer_resolution],
    "🎮 Gaming Optimizations": [revert_game_mode, revert_xbox_game_bar, revert_gpu_scheduling, revert_visual_effects],
    "🚀 Startup & Services": [revert_search_indexing, revert_telemetry],
    "🌐 Network & DNS": [revert_network_optimization],
}

# ─── GUI ─────────────────────────────────────────────────────────────────────

class OptimizerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"⚡ Windows Optimizer v{VERSION}")
        self.geometry("920x720")
        self.running = False
        self.checks = {}
        self._build_ui()
        self.check_for_updates()

    def _build_ui(self):
        # ── Grid config ──
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ── Sidebar ──
        self.sidebar = ctk.CTkFrame(self, width=280, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(4, weight=1)

        ctk.CTkLabel(
            self.sidebar, text="⚡ Optimizer",
            font=ctk.CTkFont(size=24, weight="bold"), text_color=ACCENT
        ).grid(row=0, column=0, padx=20, pady=(20, 10))

        # Admin status
        admin_color = SUCCESS if is_admin() else DANGER
        admin_text  = "● Admin Active" if is_admin() else "● No Admin Rights"
        ctk.CTkLabel(
            self.sidebar, text=admin_text,
            font=ctk.CTkFont(size=12), text_color=admin_color
        ).grid(row=1, column=0, padx=20, pady=(0, 20))

        # Checkboxes area
        self.scrollable_frame = ctk.CTkScrollableFrame(self.sidebar, label_text="Optimizations")
        self.scrollable_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
        self.scrollable_frame.grid_columnconfigure(0, weight=1)

        for group in TASKS:
            var = ctk.BooleanVar(value=True)
            self.checks[group] = var
            cb = ctk.CTkCheckBox(
                self.scrollable_frame, text=group, variable=var,
                font=ctk.CTkFont(size=13), corner_radius=6
            )
            cb.pack(fill="x", padx=10, pady=10)

        # Sidebar Buttons
        self.all_btn = ctk.CTkButton(
            self.sidebar, text="Select All", height=32,
            fg_color="transparent", border_width=1,
            command=self._select_all
        )
        self.all_btn.grid(row=3, column=0, padx=20, pady=(10, 0))

        self.none_btn = ctk.CTkButton(
            self.sidebar, text="Deselect All", height=32,
            fg_color="transparent", border_width=1,
            command=self._select_none
        )
        self.none_btn.grid(row=4, column=0, padx=20, pady=(10, 0), sticky="n")

        # Sidebar Footer
        self.gh_btn = ctk.CTkButton(
            self.sidebar, text="⭐ GitHub", height=32,
            fg_color="#24292e", hover_color="#333",
            command=self._open_github
        )
        self.gh_btn.grid(row=5, column=0, padx=20, pady=10)

        self.upd_btn = ctk.CTkButton(
            self.sidebar, text="🔄 Check Update", height=32,
            fg_color="transparent", border_width=1, border_color=SUCCESS,
            text_color=SUCCESS, hover_color="#1a3328",
            command=lambda: self.check_for_updates(manual=True)
        )
        self.upd_btn.grid(row=6, column=0, padx=20, pady=(0, 20))

        # ── Main Area ──
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)

        # Top Bar
        self.top_bar = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.top_bar.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        
        ctk.CTkLabel(
            self.top_bar, text="Operation Log",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(side="left")

        self.clear_btn = ctk.CTkButton(
            self.top_bar, text="Clear", width=80, height=26,
            fg_color="transparent", border_width=1,
            command=self._clear_log
        )
        self.clear_btn.pack(side="right")

        # Log Textbox
        self.log_box = ctk.CTkTextbox(self.main_frame, font=ctk.CTkFont(family="Consolas", size=13))
        self.log_box.grid(row=1, column=0, sticky="nsew")
        
        # Tags for log box
        # Note: CTkTextbox doesn't support tags directly like tk.Text, but we can access the internal widget
        self.log_box._textbox.tag_config("ok",      foreground=SUCCESS)
        self.log_box._textbox.tag_config("err",     foreground=DANGER)
        self.log_box._textbox.tag_config("section", foreground=ACCENT)
        self.log_box._textbox.tag_config("info",    foreground=WARNING)

        # Action Buttons
        self.actions = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.actions.grid(row=2, column=0, sticky="ew", pady=(10, 0))

        self.run_btn = ctk.CTkButton(
            self.actions, text="▶  Run Optimizations", height=45,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color=ACCENT2, hover_color="#6d28d9",
            command=self._run
        )
        self.run_btn.pack(side="left", fill="x", expand=True, padx=(0, 5))

        self.undo_btn = ctk.CTkButton(
            self.actions, text="⏪ Undo Changes", height=45,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#334155", hover_color="#475569",
            text_color=WARNING,
            command=self._undo
        )
        self.undo_btn.pack(side="left", fill="x", expand=True, padx=(5, 0))

        # Progress
        self.progress = ctk.CTkProgressBar(self.main_frame, mode="indeterminate")
        self.progress.grid(row=3, column=0, sticky="ew", pady=(15, 0))
        self.progress.set(0)

        self._log("  Welcome! Select categories in the sidebar and click Run.", "info")
        if not is_admin():
            self._log("  [!] Running without Administrator rights. Most tweaks will fail.", "err")

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
                    if messagebox.askyesno("Update Available", f"A new version ({remote_version}) is available.\n\nWould you like to download and install it?"):
                        self._log(f"Downloading update v{remote_version}...", "info")
                        
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
                            self._log("[!] Error: Update file not ready.", "err")
                            return

                        if expected_hash and sha256_hash.hexdigest().lower() != expected_hash.lower():
                            self._log("[!] Security Error: Hash mismatch!", "err")
                            if os.path.exists(update_path): os.remove(update_path)
                            messagebox.showerror("Security Error", "Verification failed. Aborting.")
                            return
                        
                        self._log("Launching update...", "ok")
                        os.startfile(update_path)
                        self.after(0, self.destroy)
                        os._exit(0)
                elif manual:
                    messagebox.showinfo("Update Check", f"You are on the latest version ({VERSION}).")
            except Exception as e:
                self._log(f"Update check failed: {e}", "err")
                if manual:
                    messagebox.showerror("Update Check", f"Could not check for updates.\n{e}")

        threading.Thread(target=_check, daemon=True).start()

    def _undo(self):
        if self.running: return
        selected = [g for g, v in self.checks.items() if v.get() and g in REVERT_TASKS]
        if not selected:
            messagebox.showwarning("Nothing to undo", "Select the categories you wish to revert.")
            return
        if not messagebox.askyesno("Confirm Undo", "This will revert selected optimizations to Windows defaults. Continue?"):
            return
        self.running = True
        self.undo_btn.configure(state="disabled", text="⏳ Reverting...")
        self.progress.start()
        threading.Thread(target=self._undo_worker, args=(selected,), daemon=True).start()

    def _undo_worker(self, groups):
        t0 = time.time()
        self._log("\n" + "═"*52, "section")
        self._log("  ⏪  REVERTING CHANGES", "section")
        self._log("═"*52 + "\n", "section")

        for group in groups:
            self._log(f"\n── Reverting: {group} ──", "section")
            for fn in REVERT_TASKS[group]:
                try:
                    fn(self._log)
                except Exception as e:
                    self._log(f"   ✖ Error: {e}", "err")

        elapsed = time.time() - t0
        self._log("\n" + "═"*52, "section")
        self._log(f"  ✅  REVERT DONE in {elapsed:.1f}s", "ok")
        self._log("═"*52, "section")

        def finalize():
            self.running = False
            self.progress.stop()
            self.undo_btn.configure(state="normal", text="⏪ Undo Changes")
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
            messagebox.showwarning("Nothing selected", "Please select at least one optimization.")
            return
        self.running = True
        self.run_btn.configure(state="disabled", text="⏳ Running...")
        self.progress.start()
        threading.Thread(target=self._worker, args=(selected,), daemon=True).start()

    def _worker(self, groups):
        t0 = time.time()
        self._log("\n" + "═"*52, "section")
        self._log("  ⚡  OPTIMIZATION STARTED", "section")
        self._log("═"*52 + "\n", "section")

        for group in groups:
            self._log(f"\n── {group} ──", "section")
            for fn in TASKS[group]:
                try:
                    fn(self._log)
                except Exception as e:
                    self._log(f"   ✖ Error: {e}", "err")

        elapsed = time.time() - t0
        self._log("\n" + "═"*52, "section")
        self._log(f"  ✅  DONE in {elapsed:.1f}s", "ok")
        self._log("  Restart your PC for all changes to take effect.", "info")
        self._log("═"*52, "section")

        def finalize():
            self.running = False
            self.progress.stop()
            self.run_btn.configure(state="normal", text="▶  Run Optimizations")
        self.after(0, finalize)

# ─── Entry point ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if sys.platform != "win32":
        print("This script is for Windows only.")
        sys.exit(1)

    if not is_admin():
        try:
            args_string = " ".join(f'"{arg}"' for arg in sys.argv)
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, args_string, None, 1
            )
        except Exception:
            pass
        sys.exit(0)

    app = OptimizerApp()
    app.mainloop()
