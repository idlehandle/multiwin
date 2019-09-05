import psutil
import datetime
import tkinter as tk
import pywinauto as pyw
from tkinter import messagebox

# TODO:
# Features to add:
#   - override default, smaller window
#   - Smaller font for ALL display
#   - add logger
#   - add config to save lock statuses
#   - Some window management - always on top (close and pop up added)


class Window:
    # Store all the meaningful window and associated tk info in one convenient location

    maximized_state = {
        True: 'Maximized',
        False: 'Windowed'
    }

    def __init__(self, hwnd, master):
        # Basic info
        self.master = master
        self.hwnd = hwnd
        self.name = self.info().name
        try:
            self.short_name = (self.name[:64] + ' ...') if len(self.name) > 64 else self.name
        except TypeError:
            self.short_name = self.name
        self.process_id = self.wrapper().process_id()
        self.process_name = psutil.Process(self.process_id).name()
        self._is_resetting = False

        # Variables
        self.is_locked = tk.BooleanVar()
        self.is_locked.set(False)
        self.is_locked.trace_add('write', self._lock_functions)
        self.is_max_text_var = tk.StringVar()
        self.is_maximized = tk.BooleanVar()
        self.is_maximized.trace_add('write', self._btn_maximize_trace)
        self.is_in_position = tk.IntVar()
        self.is_in_position.trace_add('write', self._rb_pos_change)

    def create_lock_button(self, process_group):
        self._is_resetting = True
        self.btn_lock = tk.Checkbutton(
            process_group,
            text=self.short_name,
            variable=self.is_locked,
            indicatoron=0,
            justify=tk.LEFT,
            anchor=tk.W,
            relief=tk.GROOVE,
            selectcolor='bisque'
        )
        self.btn_lock.bind('<Button-2>', lambda evt: self.set_focus_keep_cursor())
        self.btn_lock.bind('<Button-3>', self.binding_close_window)
        self.btn_lock.bind('<Enter>', lambda evt, win=self.name: self.master._change_hover_text(win))
        self.btn_lock.bind('<Leave>', lambda evt: self.master._change_hover_text(''))
        self._is_resetting = False

    def create_maximize_button(self, process_group):
        self._is_resetting = True
        maxed = self.wrapper().is_maximized()
        self.is_maximized.set(maxed)
        self.is_max_text_var.set(Window.maximized_state.get(maxed))
        self.btn_maximize = tk.Checkbutton(
            process_group,
            textvariable=self.is_max_text_var,
            variable=self.is_maximized,
            indicatoron=0,
            relief=tk.FLAT,
            selectcolor='gold'
        )
        self._is_resetting = False

    def create_radio_buttons(self, process_group, monitors):
        self._is_resetting = True
        self.is_in_position.set(self.where()[0])
        self.rb_displays = [
            tk.Radiobutton(
                process_group,
                text=mon.name,
                variable=self.is_in_position,
                value=idx,
                indicatoron=0,
                relief=tk.FLAT,
                selectcolor='dodgerblue',    # 'systemHighlight',
                **GUI.pads
            ) for idx, mon in enumerate(monitors)
        ]
        for idx, radio in enumerate(self.rb_displays):
            radio.bind("<Button-3>", lambda e, i=idx: self.binding_click_all_display(e, self.process_name, i))
        self._is_resetting = False

    def binding_click_all_display(self, event, group, idx):
        # Function to change all display to one group
        # TODO: see if possible to use windows class only
        for window in self.master.process_groups.get(group):
            window.rb_displays[idx].invoke()

    def binding_close_window(self, event):
        if messagebox.askyesnocancel('Close Window', f'Close this window?\n{self.name}'):
            try:
                self.wrapper().close()
            except Exception as e:
                print(e)
            self.master.refresh(manual_call=True)

    def where(self):
        # Find min/max x value
        x = self.wrapper().rectangle().mid_point().x
        if x < Monitor.min_x:
            x = Monitor.min_x
        elif x > Monitor.max_x:
            x = Monitor.max_x

        # Check through Monitors
        for i, monitor in enumerate(self.master.monitors):
            if x in monitor.x_range:
                offset_x = (self.info().rectangle.left - monitor.x) / len(monitor.x_range)
                offset_y = (self.info().rectangle.top - monitor.y) / len(monitor.y_range)
                return i, monitor, offset_x, offset_y
        return -1, None, 0, 0

    def _lock_functions(self, *args):
        # Disable row when item is locked
        btn_state = tk.DISABLED if self.is_locked.get() else tk.NORMAL
        self.btn_maximize.configure(state=btn_state)
        for rb in self.rb_displays:
            rb.configure(state=btn_state)

    def _btn_maximize_trace(self, *args):
        # Tracer for maximize variable change
        maxed = self.wrapper().is_maximized()
        if not self._is_resetting:
            if maxed:
                self.wrapper().restore()
            else:
                self.wrapper().maximize()
            self.is_max_text_var.set(Window.maximized_state.get(not maxed))
            self.get_current_value()
        else:
            self.is_max_text_var.set(Window.maximized_state.get(maxed))

    def _rb_pos_change(self, *args):
        if not self._is_resetting:
            cur_pos = self.where()
            pos = self.is_in_position.get()
            maxed = self.wrapper().is_maximized()
            if not pos == cur_pos[0]:
                if maxed:
                    self.wrapper().restore()
                monitor = GUI.monitors[pos]
                self.wrapper().move_window(
                    monitor.x + round(len(monitor.x_range) * cur_pos[2]),
                    monitor.y + round(len(monitor.y_range) * cur_pos[3])
                )
                if maxed:
                    self.wrapper().maximize()
                self.set_focus_keep_cursor()
            self.get_current_value()

    def set_focus_keep_cursor(self):
        cursor_pos = pyw.win32api.GetCursorPos()
        self.wrapper().set_focus()
        pyw.win32api.SetCursorPos(cursor_pos)

    def get_current_value(self):
        # Refresh current maximize status and position display
        self._is_resetting = True
        self.is_maximized.set(self.wrapper().is_maximized())
        self.is_in_position.set(self.where()[0])
        self._is_resetting = False

    def info(self):
        return pyw.controls.hwndwrapper.HwndElementInfo(self.hwnd)

    def wrapper(self):
        return pyw.controls.hwndwrapper.HwndWrapper(self.hwnd)


class Monitor:
    # Meaningful data for Monitor from the windows object
    min_x = 0
    max_x = 0

    def __init__(self, hwnd):
        self.info = pyw.win32api.GetMonitorInfo(hwnd[0])
        self.name = self.info.get('Device').strip(r'\.')
        self.x, self.y, self.w, self.h = self.info.get('Monitor')
        self.x_range = range(self.x, self.w)
        self.y_range = range(self.y, self.h)
        Monitor.min_x = min(Monitor.min_x, self.x)
        Monitor.max_x = max(Monitor.max_x, self.w)


# main GUI
class GUI(tk.Tk):

    # Universal padding
    pads = {
        'padx': 2,
        'pady': 2
    }
    monitors = sorted([Monitor(mon) for mon in pyw.win32api.EnumDisplayMonitors()], key=lambda m: m.x)
    default_min = 60000
    preset_freqs = [
        ('2 min', 2 * default_min),
        ('5 min', 5 * default_min),
        ('15 min', 15 * default_min),
        ('30 min', 30 * default_min),
        ('60 min', 60 * default_min),
        ('Never', -1)
    ]
    # is_perpetuating = True

    def __init__(self):
        # Initial settings
        super().__init__()
        self.title("Multi Window Manager")
        self._job = None
        self._update_job = None
        self.protocol("WM_DELETE_WINDOW", self._seek_and_destroy)
        self.windows = dict()

        # tk variables and trace bindings
        self.freq_index = tk.IntVar()
        self.freq_index.trace_add('write', self._update_freq)
        self.freq_index.set(2)  # default frequency pair
        self.stay_on_top = tk.BooleanVar()
        self.stay_on_top.set(False)
        self.stay_on_top.trace_add('write', callback=self._update_topmost)
        self.status = tk.StringVar()
        self.window_expanded_name = ''

        # Set up menu bars
        menu_bar = tk.Menu(self)
        freq_bar = tk.Menu(menu_bar, tearoff=0)
        win_bar = tk.Menu(menu_bar, tearoff=0)

        # Add preset frequencies
        for idx, (name, freq) in enumerate(GUI.preset_freqs):
            freq_bar.add_radiobutton(
                label=name,
                value=idx,
                variable=self.freq_index,
                command=lambda i=idx, f=freq: self.set_freq(i, f)
            )
        win_bar.add_checkbutton(label='Stay atop', underline=0, variable=self.stay_on_top)
        win_bar.add_command(label='Tiny', underline=0, command=self.iconify)
        win_bar.add_command(label='Exit', underline=1, command=self._seek_and_destroy)

        menu_bar.add_cascade(label='Tool', underline=0, menu=win_bar)
        menu_bar.add_cascade(label='Set Refresh Frequency', underline=12, menu=freq_bar)
        menu_bar.add_command(label='Refresh', underline=0, command=lambda: self.refresh(True))
        self.configure(menu=menu_bar)

        # keyboard shortcut bindings
        self.bind('<<Alt_L-R>>', lambda: self.refresh(True))
        self.bind('<<Alt_L-S>>', lambda: self.stay_on_top.set(not self.stay_on_top.get()))
        self.bind('<<Alt_L-X>>', self._seek_and_destroy)
        self.bind('<<Alt_L-T>>', self.iconify)

        # set up status bar
        self._pid = psutil.Process().pid
        self.status_bar = tk.Label(
            self,
            textvariable=self.status,
            relief=tk.GROOVE,
            justify=tk.RIGHT,
            foreground='systemDisabledText',
            anchor=tk.E
        )
        self.status_bar.pack(expand=True, fill=tk.X, side=tk.BOTTOM)
        self.refresh(manual_call=True)
        self._update_status()

    def _update_topmost(self, name, index, operation):
        self.wm_attributes('-topmost', self.stay_on_top.get())

    def _seek_and_destroy(self):
        if self._update_job is not None:
            self.after_cancel(self._update_job)
            self._update_job = None
        if self._job is not None:
            self.after_cancel(self._job)
            self._job = None
        self.destroy()

    def _change_hover_text(self, value):
        self.window_expanded_name = value
        if self.window_expanded_name:
            self.status_bar.configure(background='light goldenrod')
            self.after_cancel(self._update_job)
            self._update_job = None
            self.status.set(self.window_expanded_name)
            self.status_bar.configure(anchor=tk.W)
        else:
            self.status_bar.configure(background='SystemButtonFace')
            self.status_bar.configure(anchor=tk.E)
            if self._update_job is None:
                self._update_status()

    def _update_status(self, delay=1000):
        pid = psutil.Process()
        mem = pid.memory_full_info()[-1] / 2**20  # Memory in MB
        cpu = pid.cpu_percent()
        status_string = f'Next refresh: {self.next_refresh} | Last refreshed: {self.last_refreshed} | cpu: {cpu} % | memory: {mem:,.2f} MB '
        self.status.set(status_string)
        self._update_job = self.after(ms=delay, func=self._update_status)

    def set_freq(self, idx, freq):
        self.freq_index.set(idx)
        self.next_refresh = datetime.datetime.strftime(datetime.datetime.now() + datetime.timedelta(seconds=self.refresh_freq // 1000), '%H:%M:%S')
        self.perpetuate(reset=True)

    def _update_freq(self, *args):
        self.refresh_freq = GUI.preset_freqs[self.freq_index.get()][1]

    def perpetuate(self, reset=False):
        if reset:
            if self._job is not None:
                self.after_cancel(self._job)
                self._job = None
            if self.refresh_freq >= 0:
                self._job = self.after(ms=self.refresh_freq, func=self.perpetuate)
        else:
            self.refresh(keep_perpetuate=False)
            self._job = self.after(ms=self.refresh_freq, func=self.perpetuate)

    # Allow user to refresh the screen without closing the program
    def refresh(self, manual_call=False, keep_perpetuate=True):
        try:
            self.main_frame.destroy()
        except AttributeError as e:
            # ignore error raised when main_frame didn't exist
            print(e)
        self.main_frame = tk.Frame(self, **GUI.pads)
        self.get_windows()
        self.create_windows()
        self.main_frame.pack(expand=True, fill=tk.X, side=tk.TOP)
        self.last_refreshed = datetime.datetime.now().strftime('%H:%M:%S')
        self.next_refresh = datetime.datetime.strftime(datetime.datetime.now() + datetime.timedelta(seconds=self.refresh_freq // 1000), '%H:%M:%S')
        if keep_perpetuate:
            self.perpetuate(reset=manual_call)

    # Find all the visible windows and create Window instances
    def get_windows(self):
        wins = set(w for w in pyw.findwindows.find_windows())

        # Add new to Window instances
        self.windows = {win: self.windows.get(win, Window(win, self)) for win in wins}

    def create_windows(self):
        # process groups dictionary for display settings
        self.process_groups = dict()

        # Manage Exclusions
        exclusions = {
            'name': ('Program Manager', '', ' '),
            'process_name': ('ApplicationFrameHost.exe', 'SelfService.exe', 'SystemSettings.exe')
        }

        for row, window in enumerate(sorted(self.windows.values(), key=lambda w: (w.process_name, w.name))):

            # Exclusions
            # Don't need Program Manager and hidden services
            if any(getattr(window, attr) in cond for attr, cond in exclusions.items()):
                continue

            # group processes
            if self.process_groups.get(window.process_name) is None:
                process_group = tk.LabelFrame(
                    self.main_frame,
                    text=window.process_name,
                    **GUI.pads
                )
                for m in range(len(GUI.monitors) + 2):
                    process_group.columnconfigure(m, weight=6 if m == 0 else 1, uniform='Uniform')
                process_group.pack(expand=True, fill=tk.X)
                self.process_groups[window.process_name] = list()

            # create checkbox labels within group
            window.create_lock_button(process_group)
            window.btn_lock.grid(row=row, column=0, sticky=tk.EW)

            # create maximize/restore button
            window.create_maximize_button(process_group)
            window.btn_maximize.grid(row=row, column=1, sticky=tk.EW)

            # create radio buttons within frame
            window.create_radio_buttons(process_group, GUI.monitors)
            for idx, rb in enumerate(window.rb_displays):
                rb.grid(row=row, column=idx + 2, sticky=tk.E)

            # add window to process group
            self.process_groups[window.process_name].append(window)

            # refresh with latest data
            window.get_current_value()

    # def set_all_display(self, group, idx):
    #     # function to change all display to one group
    #     for window in self.process_groups.get(group):
    #         window.rb_displays[idx].invoke()

if __name__ == '__main__':
    gui = GUI()
    gui.mainloop()
