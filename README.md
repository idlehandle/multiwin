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
6. Proportional movement between displays, e.g. if window is near top left of `DISPLAY1`, it will move to top left of `DISPLAY2`
7. Option to set transparency and always on top
8. Auto hide/re-appear when window gains/loses focus
9. Exclusion management in a `json` file

# Usage
1. Left click on Windows to lock the maximize/display buttons
2. Mid click on Windows to pop them to foreground
3. Right click to close (some windows might require saving)
4. Left click on individual displays to move the associated window to the display
5. Right click on individual display to move ALL windows in the same process (e.g. Excel.exe) to the same display
6. Hover on top of Windows name to see full name in status bar
*Exclusion management:*
7. Control Right Click on window button to exclude the process name
8. Shift Right Click on window button to exclude any matches of the name
9. Control-Shift Right Click on window button to exclude the specific combination of name and process

# Caveats (Known imperfections)
- Python itself crashes after the tool is exited.  I suspect this has to do with `tkinter` not being disposed properly in the background but need to dive further down as the error isn't caught inside Python, but post execution.
  - This is in relation to the combination of `tkinter` + `pywinauto`.  Submitted issue under `pywinauto`: https://github.com/pywinauto/pywinauto/issues/813
- Some temp/back processes still get picked up by the tool, these needs to be weeded out in `exclusions` condition
- Currently only support Windows, no plan for other OS at the moment
- Currently only support horizontally aligned similar displays, not yet tested on vertical displays with varying resolution/orientation (but it *should* work)

# Future enhancements (TBD)
- add logger
- immigrate other GUI presets into config file, track locked window status
*Nice to haves*
- Some extra window management - toggle always on top for external windows
- override default to create smaller UI
- process monitor to see Mem/CPU usage on processes
