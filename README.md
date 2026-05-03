![icon](WindowsOptimizer.ico)
# ⚡ Windows Performance Optimizer

![Version](https://img.shields.io/badge/version-1.3.4-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows-lightgray.svg)
![Python](https://img.shields.io/badge/python-3.8%2B-green.svg)

A lightweight, open-source Windows optimization utility written in Python. This tool applies system-level tweaks to reduce background bloat, lower latency, and improve frame rates for gaming and heavy workloads.

## Features

- **🗑 Deep Clean:** Clears Temp files, Prefetch, Recycle Bin, and performs a native DISM Component Store cleanup.
- **⚡ Power & CPU:** Forces Ultimate Performance power plans and adjusts Win32 timer resolution for foreground tasks.
- **🎮 Gaming Tweaks:** Enables Game Mode & Hardware-Accelerated GPU Scheduling (HAGS) while disabling Xbox DVR overlay.
- **🚀 Bloatware & Telemetry:** Disables common startup apps, telemetry data collection, and Windows Search indexing.
- **🌐 Network Tuning:** Flushes DNS and optimizes the TCP network stack for lower latency gaming.

## Installation & Usage

Because this script relies strictly on Python's built-in standard libraries, there are no external dependencies (`pip install`) required.

1. Clone the repository:
   ```bash
   git clone [https://github.com/TR4IS/WindowsOptimizer.git](https://github.com/TR4IS/WindowsOptimizer.git)
   cd WindowsOptimizer
   ```
2. install the dependencies:
   ```bash
   pip install customtkinter
   ```
2. Run the script:

   ```Bash
   python WindowsOptimizer.py
   ```

   
Note: The application will automatically prompt for Administrator privileges via UAC if not already elevated.

## ⚠ Disclaimer 

**Requires Administrator** — most registry and service changes won't work without it. The script auto-requests elevation.

**Restart needed** — HAGS, power plan, and some registry changes only take full effect after a reboot.

**DISM /resetbase is irreversible** — users can't roll back Windows updates after that step. Worth a warning in your description.

**Disabling Search indexing affects Windows Search speed** — some users may not want this.

**Startup program removal is based on a hardcoded list** — it won't catch every bloat app on every system.

**No undo/restore** — the script doesn't save previous settings. Consider adding a restore/backup feature before publishing if you want to be safe for less technical users.
