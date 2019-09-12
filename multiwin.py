import json
import psutil
import datetime
import tkinter as tk
import pywinauto as pyw
from os import path
from tkinter import messagebox, simpledialog


# TODO:
# Features to add:
#   - Smaller font for ALL display
#   - add logger
#   - Some window management - always on top (close and pop up added)
#   - DONE: move some GUI configs to config file
#   - DONE: override default, smaller window
#   - DONE: add config to save lock statuses
#   - DONE: add exclusion manager, possibly using json.  Explore in-app editing


# Store all the meaningful window and associated tk info in one convenient location
class Window:

    maximized_state = {
        True: 'Maximized',
        False: 'Windowed'
    }
    delim = chr(215) + chr(247) + chr(215)

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
        self.name_with_process = Window.delim.join((str(self.process_name), str(self.name)))
        self._is_resetting = False

        # Variables
        self.is_locked = tk.BooleanVar()
        self.is_locked.set(self.name_with_process in self.master.cfg.locked_windows)
        self.is_locked.trace_add('write', self._lock_functions)
        self.is_max_text_var = tk.StringVar()
        self.is_maximized = tk.BooleanVar()
        self.is_maximized.trace_add('write', self._btn_maximize_trace)
        self.is_in_position = tk.IntVar()
        self.is_in_position.trace_add('write', self._rb_pos_change)

    # Surpress function so additional tracers are not triggered during refresh
    def __surpressed__(function):
        def surpressed_func(self, *args, **kwargs):
            self._is_resetting = True
            function(self, *args, **kwargs)
            self._is_resetting = False
        return surpressed_func

    def get_locked_state(self):
        return tk.DISABLED if self.is_locked.get() else tk.NORMAL

    @__surpressed__
    def create_lock_button(self, process_group):
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
        self.btn_lock.bind('<Button-2>', lambda evt: self.anchored_cursor_wrapper(self.wrapper().set_focus))
        self.btn_lock.bind('<Button-3>', self.binding_close_window)
        self.btn_lock.bind('<Control-Shift-Button-3>', lambda evt, ex_win=self, ex_typ='name_with_process': self.master.cfg.exclude_item(ex_typ, ex_win))
        self.btn_lock.bind('<Control-Button-3>', lambda evt, ex_win=self, ex_typ='process_name': self.master.cfg.exclude_item(ex_typ, ex_win))
        self.btn_lock.bind('<Shift-Button-3>', lambda evt, ex_win=self, ex_typ='name': self.master.cfg.exclude_item(ex_typ, ex_win))
        self.btn_lock.bind('<Enter>', lambda evt, win=self.name: self.master._change_hover_text(win))
        self.btn_lock.bind('<Leave>', lambda evt: self.master._change_hover_text(''))

    @__surpressed__
    def create_maximize_button(self, process_group):
        maxed = self.wrapper().is_maximized()
        self.is_maximized.set(maxed)
        self.is_max_text_var.set(Window.maximized_state.get(maxed))
        self.btn_maximize = tk.Checkbutton(
            process_group,
            textvariable=self.is_max_text_var,
            variable=self.is_maximized,
            indicatoron=0,
            relief=tk.FLAT,
            selectcolor='gold',
            state=self.get_locked_state()
        )

    @__surpressed__
    def create_radio_buttons(self, process_group, monitors):
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
                state=self.get_locked_state(),
                **GUI.pads
            ) for idx, mon in enumerate(monitors)
        ]
        for idx, radio in enumerate(self.rb_displays):
            radio.bind("<Button-3>", lambda e, i=idx: self.binding_click_all_display(e, self.process_name, i))

    def binding_click_all_display(self, event, group, idx):
        # Function to change all display to one group
        # TODO: see if possible to use windows class only
        for window in self.master.process_groups.get(group):
            window.rb_displays[idx].invoke()

    def binding_close_window(self, event):
        if messagebox.askyesnocancel('Close Window', f'Close this window?\n{self.name}'):
            try:
                self.anchored_cursor_wrapper(self.wrapper().close_alt_f4)
            except Exception as e:
                print(e)
            self.master.refresh(manual_call=True)

    # Find out where the window is currently located
    def where(self):
        # Find min/max x value
        x = self.wrapper().rectangle().mid_point().x
        if x < Monitor.min_x:
            x = Monitor.min_x
        elif x > Monitor.max_x:
            x = Monitor.max_x

        # Check through Monitors
        # returns (index_in_monitor_list, monitor_object, offset_x, offset_y)
        for i, monitor in enumerate(self.master.monitors):
            if x in monitor.x_range:
                offset_x = (self.info().rectangle.left - monitor.x) / len(monitor.x_range)
                offset_y = (self.info().rectangle.top - monitor.y) / len(monitor.y_range)
                return i, monitor, offset_x, offset_y
        return -1, None, 0, 0

    def _lock_functions(self, *args):
        # Disable row when item is locked
        if self.is_locked.get():
            self.master.cfg.locked_windows.append(self.name_with_process)
        else:
            self.master.cfg.locked_windows[:] = [win for win in self.master.cfg.locked_windows if win != self.name_with_process]
        btn_state = self.get_locked_state()
        self.btn_maximize.configure(state=btn_state)
        for rb in self.rb_displays:
            rb.configure(state=btn_state)

    # Tracer for maximize variable change
    def _btn_maximize_trace(self, *args):
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

    # Tracing for radio button position change
    def _rb_pos_change(self, *args):
        if not self._is_resetting:

            # First, check if window is being moved
            cur_pos = self.where()
            pos = self.is_in_position.get()
            maxed = self.wrapper().is_maximized()
            if not pos == cur_pos[0]:
                # Window is in a different display
                # Restore and re-maximize window if they are already maximized

                if maxed:
                    self.wrapper().restore()
                monitor = GUI.monitors[pos]
                self.wrapper().move_window(
                    monitor.x + round(len(monitor.x_range) * cur_pos[2]),
                    monitor.y + round(len(monitor.y_range) * cur_pos[3])
                )
                if maxed:
                    self.wrapper().maximize()
                self.anchored_cursor_wrapper(self.wrapper().set_focus)
                self.get_current_value()

    # Keep cursor in same position while function is executed
    def anchored_cursor_wrapper(self, func):
        cursor_pos = pyw.win32api.GetCursorPos()
        func()
        pyw.win32api.SetCursorPos(cursor_pos)

    # Refresh current maximize status and position display
    @__surpressed__
    def get_current_value(self):
        self.is_maximized.set(self.wrapper().is_maximized())
        self.is_in_position.set(self.where()[0])

    def info(self):
        return pyw.controls.hwndwrapper.HwndElementInfo(self.hwnd)

    def wrapper(self):
        return pyw.controls.hwndwrapper.HwndWrapper(self.hwnd)


# Translate meaningful data for Monitor from the windows object
class Monitor:
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


class Config:
    # Class for advance management of config and exclusions

    default_config = {
        'transparency': {
            'Min': 0.1,
            'Inactive': 0.2,
            'Active': 0.8,
            'Max': 1.0
        },
        'exclusions': {
            'name': [''],
            'name_with_process': [],
            'process_name': ['ApplicationFrameHost.exe', 'SelfService.exe', 'SystemSettings.exe'],
        },
        'locked_windows': [],
        'freq_index': 2,
        'stay_on_top': False
    }

    def __init__(self, excl: dict=None):
        self.config_path = path.dirname(path.abspath(__file__))
        self.config_file = path.join(self.config_path, 'multi_config.json')
        self.load()
        if excl:
            self.exclusions.update(excl)

    # Load and set default if file not located
    def load(self):
        try:
            with open(self.config_file, 'r+') as file:
                self.data = json.load(file)
        except FileNotFoundError:
            self.data = Config.default_config
        for attr, val in self.data.items():
            setattr(self, attr, val)

    def save(self):
        with open(self.config_file, 'w+') as file:
            json.dump(self.data, file, indent=2)

    # Pop up window to ask if users want to exclude the window / process / both
    def exclude_item(self, exclude_type, exclude_window):
        exclude_descriptions = {
            'name': f'Hide all Windows with this name?\n\n{exclude_window.name}',
            'process_name': f'Hide all Windows under this process type?\n\n{exclude_window.process_name}',
            'name_with_process': f'Hide this particular Window?\n\n{exclude_window.name}\nUnder: {exclude_window.process_name}'
        }
        if messagebox.askyesnocancel(
            'Add Exclusion',
            exclude_descriptions.get(exclude_type)
        ):
            self.add_exclusion(exclude_type, getattr(exclude_window, exclude_type))
            exclude_window.master.refresh(manual_call=True)

    def add_exclusion(self, attr: str, value: str):
        excl_list = self.exclusions.get(attr, None)
        if excl_list is not None:
            excl_list.append(value)

    def is_excluded(self, window):
        if self.exclusions:
            return any(getattr(window, attr, '!No Attribute!') in cond for attr, cond in self.exclusions.items())
        else:
            return False

    def update(self, item_and_value: dict):
        self.data.update(item_and_value)


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
    mini_resolution = '120x60'
    event_threshold = default_min // 2

    # Initial settings
    def __init__(self):
        super().__init__()
        self.title("Multi Window Manager")
        self._job_perpetuate = None
        self._job_update = None
        self._job_minimize = None
        self._mouse_left = datetime.datetime.now()

        # Load config for Transparency, Exclusions...
        self.cfg = Config()
        self.protocol("WM_DELETE_WINDOW", self._exit_strategy)
        self.windows = dict()
        self.transparency_order = sorted(self.cfg.transparency.keys(), key=lambda k: self.cfg.transparency.get(k))

        # tk variables and trace bindings
        self.freq_index = tk.IntVar()
        self.freq_index.trace_add('write', self._update_freq)
        self.freq_index.set(self.cfg.freq_index)
        self.stay_on_top = tk.BooleanVar()
        self.stay_on_top.set(self.cfg.stay_on_top)
        self.stay_on_top.trace_add('write', callback=self._update_topmost)
        self.status = tk.StringVar()
        self.window_hover_expanded_name = ''

        # Set up menu bars
        menu_bar = tk.Menu(self)
        freq_bar = tk.Menu(menu_bar, tearoff=0)
        win_bar = tk.Menu(menu_bar, tearoff=0)
        trans_bar = tk.Menu(menu_bar, tearoff=0)

        # Add preset frequencies
        for idx, (name, freq) in enumerate(GUI.preset_freqs):
            freq_bar.add_radiobutton(
                label=name,
                value=idx,
                variable=self.freq_index,
                command=lambda i=idx, f=freq: self.set_freq(i, f)
            )
        for trans in self.transparency_order[:0:-1]:
            trans_bar.add_command(label=f'Set {trans} Transparency', command=lambda t=trans: self._set_transparency(t))

        win_bar.add_checkbutton(label='Stay atop', underline=0, variable=self.stay_on_top)
        win_bar.add_cascade(label='Set Transparency', menu=trans_bar)
        win_bar.add_command(label='Tiny', underline=0, command=self.iconify)
        win_bar.add_command(label='Exit', underline=1, command=self._exit_strategy)

        menu_bar.add_cascade(label='Tool', underline=0, menu=win_bar)
        menu_bar.add_cascade(label='Set Refresh Frequency', underline=12, menu=freq_bar)
        menu_bar.add_command(label='Refresh', underline=0, command=lambda: self.refresh(True))
        self.configure(menu=menu_bar)

        # keyboard shortcut bindings
        self.bind('<<Alt_L-R>>', lambda: self.refresh(True))
        self.bind('<<Alt_L-S>>', lambda: self.stay_on_top.set(not self.stay_on_top.get()))
        self.bind('<<Alt_L-X>>', self._exit_strategy)
        self.bind('<<Alt_L-T>>', self.iconify)
        self.bind('<FocusIn>', self._got_focus)
        self.bind('<FocusOut>', self._lost_focus)
        self.bind('<Enter>', self._mouse_enter)
        self.bind('<Leave>', self._mouse_leave)

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

        # initial start
        self.refresh(manual_call=True)
        self._update_status()

    # Decorator class for binding check
    def __toplevel_check(function):
        def level_check(self, event):
            if event.widget == self:
                function(self, event)
        return level_check

    @__toplevel_check
    def _got_focus(self, event):
        if self.stay_on_top.get():
            self.wm_attributes('-alpha', self.cfg.transparency.get('Active'))
        else:
            self.wm_attributes('-alpha', self.cfg.transparency.get('Max'))

    @__toplevel_check
    def _lost_focus(self, event):
        if self.stay_on_top.get():
            self.wm_attributes('-alpha', self.cfg.transparency.get('Inactive'))
        else:
            self.wm_attributes('-alpha', self.cfg.transparency.get('Active'))

    @__toplevel_check
    def _mouse_enter(self, event):
        if self._job_minimize is not None:
            self.after_cancel(self._job_minimize)
            self._job_minimize = None
        if (datetime.datetime.now() - self._mouse_left) >= datetime.timedelta(seconds=GUI.event_threshold // 1000):
            self.geometry('')
            self.refresh(manual_call=True)

    @__toplevel_check
    def _mouse_leave(self, event):
        if self.wm_attributes('-topmost'):
            self._job_minimize = self.after(ms=GUI.event_threshold, func=lambda: self.geometry(GUI.mini_resolution))
        self._mouse_left = datetime.datetime.now()

    def _update_topmost(self, name, index, operation):
        state = self.stay_on_top.get()
        self.cfg.update({'stay_on_top': state})
        self.wm_attributes('-topmost', state)
        self.wm_attributes('-alpha', self.cfg.transparency.get('Active') if state else self.cfg.transparency.get('Max'))

    def _exit_strategy(self):
        if self._job_update is not None:
            self.after_cancel(self._job_update)
            self._job_update = None
        if self._job_perpetuate is not None:
            self.after_cancel(self._job_perpetuate)
            self._job_perpetuate = None
        self.cfg.save()
        self.destroy()

    def _change_hover_text(self, value):
        self.window_hover_expanded_name = value
        if self.window_hover_expanded_name:
            self.status_bar.configure(background='light goldenrod')
            self.after_cancel(self._job_update)
            self._job_update = None
            self.status.set(self.window_hover_expanded_name)
            self.status_bar.configure(anchor=tk.W)
        else:
            self.status_bar.configure(background='SystemButtonFace')
            self.status_bar.configure(anchor=tk.E)
            if self._job_update is None:
                self._update_status()

    def _update_status(self, delay=1000):
        pid = psutil.Process()
        mem = pid.memory_full_info()[-1] / 2**20  # Memory in MB
        cpu = pid.cpu_percent()
        status_string = f'Next refresh: {self.next_refresh} | Last refreshed: {self.last_refreshed} | cpu: {cpu} % | memory: {mem:,.2f} MB '
        self.status.set(status_string)
        self._job_update = self.after(ms=delay, func=self._update_status)

    def set_freq(self, idx, freq):
        self.cfg.update({'freq_index': idx})
        self.freq_index.set(idx)
        self.next_refresh = datetime.datetime.strftime(datetime.datetime.now() + datetime.timedelta(seconds=self.refresh_freq // 1000), '%H:%M:%S')
        self.perpetuate(reset=True)

    def _update_freq(self, *args):
        self.refresh_freq = GUI.preset_freqs[self.freq_index.get()][1]

    def _set_transparency(self, selection: str):
        index = self.transparency_order.index(selection)
        max_trans = self.cfg.transparency.get(self.transparency_order[min(index + 1, len(self.transparency_order) - 1)])
        min_trans = self.cfg.transparency.get(self.transparency_order[max(index - 1, 0)])

        transparency = simpledialog.askfloat(
            f'Set {selection} Transparency Rate',
            f'Set {selection} Transparency Rate:\r\nMinimum {min_trans}\nMaximum of {max_trans}',
            initialvalue=self.cfg.transparency.get(selection),
            minvalue=min_trans,
            maxvalue=max_trans
        )
        if transparency is not None:
            self.cfg.transparency[selection] = transparency
            if self.stay_on_top.get():
                self.wm_attributes('-alpha', self.cfg.transparency.get('Active'))
            else:
                self.wm_attributes('-alpha', self.cfg.transparency.get('Max'))

    def perpetuate(self, reset=False):
        if reset:
            if self._job_perpetuate is not None:
                self.after_cancel(self._job_perpetuate)
                self._job_perpetuate = None
            if self.refresh_freq >= 0:
                self._job_perpetuate = self.after(ms=self.refresh_freq, func=self.perpetuate)
        else:
            self.refresh(keep_perpetuate=False)
            self._job_perpetuate = self.after(ms=self.refresh_freq, func=self.perpetuate)

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
        now = datetime.datetime.now()
        self.last_refreshed = now.strftime('%H:%M:%S')
        self.next_refresh = datetime.datetime.strftime(now + datetime.timedelta(seconds=self.refresh_freq // 1000), '%H:%M:%S')
        if keep_perpetuate:
            self.perpetuate(reset=manual_call)

    # Find all the visible windows and create Window instances
    def get_windows(self, retry=0):
        if retry <= 3:
            wins = set(w for w in pyw.findwindows.find_windows())

            # Add new to Window instances
            # This try seem to slow down the process just enough to avoid the specific error it's trying to catch...
            try:
                self.windows = {win: self.windows.get(win, Window(win, self)) for win in wins}
            except pyw.controls.hwndwrapper.InvalidWindowHandle as e:
                print(e)
                print(f'retrying {retry}...')
                self.get_windows(retry + 1)

        # maybe do something if it fails after 3 times...

    def create_windows(self):
        # (Re)initiate dictionary to group windows by process name
        self.process_groups = dict()

        for row, window in enumerate(sorted(self.windows.values(), key=lambda w: (w.process_name, w.name))):

            # Exclude certain names/processes
            if self.cfg.is_excluded(window):
                continue

            # process label frame
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


if __name__ == '__main__':
    gui = GUI()
    gui.mainloop()
