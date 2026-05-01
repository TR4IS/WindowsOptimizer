"""
Windows Performance & Gaming Optimizer
Run as Administrator for full effect.
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import subprocess
import os
import sys
import ctypes
import winreg
import time
import glob
import webbrowser

# ─── App Info ────────────────────────────────────────────────────────────────
VERSION   = "1.0.0"
GITHUB_URL = "https://github.com/TR4IS/WindowsOptimizer"

# ─── Color palette ───────────────────────────────────────────────────────────
BG        = "#0d0f1a"
PANEL     = "#131627"
ACCENT    = "#00e5ff"
ACCENT2   = "#7c3aed"
SUCCESS   = "#00ff88"
WARNING   = "#ffaa00"
DANGER    = "#ff4444"
TEXT      = "#e2e8f0"
SUBTEXT   = "#64748b"
BORDER    = "#1e2540"

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

# ─── GUI ─────────────────────────────────────────────────────────────────────

class OptimizerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"⚡ Windows Performance Optimizer v{VERSION}")
        self.geometry("880x680")
        self.resizable(True, True)
        self.configure(bg=BG)
        self.running = False
        self.checks = {}
        self._build_ui()

    def _build_ui(self):
        # ── Header ──
        hdr = tk.Frame(self, bg=BG)
        hdr.pack(fill="x", padx=24, pady=(20, 0))

        tk.Label(
            hdr, text=f"⚡ Windows Optimizer v{VERSION}",
            font=("Consolas", 22, "bold"), fg=ACCENT, bg=BG
        ).pack(side="left")

        admin_color = SUCCESS if is_admin() else DANGER
        admin_text  = "● Admin" if is_admin() else "● No Admin"
        tk.Label(
            hdr, text=admin_text,
            font=("Consolas", 11), fg=admin_color, bg=BG
        ).pack(side="right", padx=4)

        if not is_admin():
            tk.Label(
                self,
                text="⚠  Run as Administrator for full effect  ⚠",
                font=("Consolas", 10), fg=WARNING, bg=BG
            ).pack(pady=(4, 0))

        # ── Divider ──
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", padx=24, pady=10)

        # ── Main layout ──
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=24, pady=0)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(0, weight=1)

        # Left: checkboxes
        left = tk.Frame(body, bg=PANEL, bd=0, highlightthickness=1,
                        highlightbackground=BORDER)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=0)

        tk.Label(
            left, text="Select Optimizations",
            font=("Consolas", 11, "bold"), fg=ACCENT, bg=PANEL, pady=10
        ).pack(fill="x", padx=12)

        tk.Frame(left, bg=BORDER, height=1).pack(fill="x", padx=12)

        for group in TASKS:
            var = tk.BooleanVar(value=True)
            self.checks[group] = var
            cb = tk.Checkbutton(
                left, text=group, variable=var,
                font=("Consolas", 10), fg=TEXT, bg=PANEL,
                selectcolor=PANEL, activebackground=PANEL,
                activeforeground=ACCENT, anchor="w",
                relief="flat", bd=0, pady=6, padx=12,
                highlightthickness=0
            )
            cb.pack(fill="x")

        # select all / none
        btn_row = tk.Frame(left, bg=PANEL)
        btn_row.pack(fill="x", padx=12, pady=8)

        tk.Button(
            btn_row, text="All", command=self._select_all,
            bg=BORDER, fg=TEXT, font=("Consolas", 9),
            relief="flat", padx=8, pady=3, cursor="hand2"
        ).pack(side="left", padx=(0, 4))

        tk.Button(
            btn_row, text="None", command=self._select_none,
            bg=BORDER, fg=TEXT, font=("Consolas", 9),
            relief="flat", padx=8, pady=3, cursor="hand2"
        ).pack(side="left")

        # Right: log
        right = tk.Frame(body, bg=PANEL, bd=0, highlightthickness=1,
                         highlightbackground=BORDER)
        right.grid(row=0, column=1, sticky="nsew")

        tk.Label(
            right, text="Output Log",
            font=("Consolas", 11, "bold"), fg=ACCENT, bg=PANEL, pady=10
        ).pack(fill="x", padx=12)

        tk.Frame(right, bg=BORDER, height=1).pack(fill="x", padx=12)

        self.log_box = scrolledtext.ScrolledText(
            right, bg=BG, fg=TEXT,
            font=("Consolas", 10), relief="flat",
            insertbackground=ACCENT, wrap="word",
            padx=12, pady=8, bd=0
        )
        self.log_box.pack(fill="both", expand=True, padx=8, pady=8)
        self.log_box.tag_config("ok",      foreground=SUCCESS)
        self.log_box.tag_config("err",     foreground=DANGER)
        self.log_box.tag_config("section", foreground=ACCENT)
        self.log_box.tag_config("info",    foreground=WARNING)

        # ── Bottom buttons ──
        bottom = tk.Frame(self, bg=BG)
        bottom.pack(fill="x", padx=24, pady=14)

        self.progress = ttk.Progressbar(bottom, mode="indeterminate", length=300)
        self.progress.pack(side="left", padx=(0, 16))

        self.run_btn = tk.Button(
            bottom, text="▶  Run Selected Optimizations",
            command=self._run,
            bg=ACCENT2, fg="white",
            font=("Consolas", 11, "bold"),
            relief="flat", padx=18, pady=8,
            cursor="hand2", activebackground="#6d28d9"
        )
        self.run_btn.pack(side="left")

        # GitHub Button
        tk.Button(
            bottom, text="⭐ GitHub",
            command=self._open_github,
            bg=BORDER, fg=ACCENT,
            font=("Consolas", 10, "bold"),
            relief="flat", padx=12, pady=8,
            cursor="hand2"
        ).pack(side="right", padx=(8, 0))

        tk.Button(
            bottom, text="Clear Log",
            command=self._clear_log,
            bg=BORDER, fg=SUBTEXT,
            font=("Consolas", 10),
            relief="flat", padx=12, pady=8,
            cursor="hand2"
        ).pack(side="right")

        self._log("  Welcome! Select the optimizations you want to apply,")
        self._log("  then click ▶ Run. Run as Administrator for best results.\n", "info")

    # ── helpers ──────────────────────────────────────────────────────────────

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
        self.progress.start(12)
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
            self.run_btn.configure(state="normal", text="▶  Run Selected Optimizations")
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