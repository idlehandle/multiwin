# multiwin
Multi window display manager

Basic python tool to manage windows opened in multiple display.  Currently only support Windows, no plans to support other OS atm.

# Requirement
Require `pywinauto` and `psutil`

# Features
1. Centralized tool for multiple monitor windows management
2. Maximize / Restore windows
3. Close / focus windows
4. Auto refresh list of windows based on preset frequencies
5. Status bar with scheduled refresh and self monitor to see memory/CPU usage

# Usage
1. Left click on Windows to lock the maximize/display buttons
2. Mid click on Windows to pop them to foreground
3. Right click to close (some windows might require saving)
4. Left click on individual displays to move the associated window to the display
5. Right click on individual display to move ALL windows in the same process (e.g. Excel.exe) to the same display

# Future enhancements (TBD)
- Proportional movement: `new_x = x // mon_x`
- override default, smaller UI
- add logger
- add config to save lock statuses
- Some extra window management - always on top
- add tooltip style hover text over window names that have been truncated (maybe in status bar)
- process monitor to see Mem/CPU usage on processes
