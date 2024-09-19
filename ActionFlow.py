import threading
import time
import json
import random
import ast
from typing import Any, Optional, List, Tuple, Dict
import keyboard
import mouse
import pyautogui
import pygetwindow as gw
import customtkinter as ctk
from pynput.keyboard import Controller, Key
from pywinauto import Application
from pywinauto.findwindows import find_windows
import win32gui
import win32con
import win32api

keyboard_controller = Controller()


KEY_NAME_MAPPING = {
    "ctrl": Key.ctrl,
    "control": Key.ctrl,
    "shift": Key.shift,
    "alt": Key.alt,
    "enter": Key.enter,
    "space": Key.space,
    "backspace": Key.backspace,
    "tab": Key.tab,
    "esc": Key.esc,
    "left": Key.left,
    "right": Key.right,
    "up": Key.up,
    "down": Key.down,
}


class CTkListbox(ctk.CTkScrollableFrame):
    def __init__(
        self,
        master: Any,
        height: int = 200,
        width: int = 300,
        highlight_color: str = "default",
        fg_color: str = "transparent",
        bg_color: Optional[str] = None,
        text_color: str = "default",
        hover_color: str = "default",
        button_color: str = "default",
        border_width: int = 3,
        font: Optional[Tuple[str, int]] = None,
        multiple_selection: bool = False,
        listvariable: Optional[ctk.StringVar] = None,
        hover: bool = True,
        command: Optional[Any] = None,
        justify: str = "left",
        **kwargs,
    ):
        super().__init__(
            master,
            width=width,
            height=height,
            fg_color=fg_color,
            border_width=border_width,
            **kwargs,
        )

        self._scrollbar.grid_configure(padx=(0, border_width + 4))
        self._scrollbar.configure(width=12)

        if bg_color:
            self.configure(bg_color=bg_color)

        theme = ctk.ThemeManager.theme
        self.select_color = (
            theme["CTkButton"]["fg_color"]
            if highlight_color == "default"
            else highlight_color
        )
        self.text_color = (
            theme["CTkLabel"]["text_color"] if text_color == "default" else text_color
        )
        self.hover_color = (
            theme["CTkButton"]["hover_color"]
            if hover_color == "default"
            else hover_color
        )

        if not font:
            self.font = ctk.CTkFont(family=theme["CTkFont"]["family"], size=13)
        else:
            if isinstance(font, ctk.CTkFont):
                self.font = font
            else:
                self.font = ctk.CTkFont(family=font[0], size=font[1])

        self.button_fg_color = (
            "transparent" if button_color == "default" else button_color
        )

        justify_map = {"left": "w", "right": "e", "center": "c"}
        self.justify = justify_map.get(justify.lower(), "w")

        self.buttons: Dict[int, ctk.CTkButton] = {}
        self.command = command
        self.multiple = multiple_selection
        self.selected: Optional[ctk.CTkButton] = None
        self.hover = hover
        self.end_num = 0
        self.selections: List[ctk.CTkButton] = []
        self.selected_index = 0
        self.columnconfigure(0, weight=1)

        if listvariable:
            self.listvariable = listvariable
            self.listvariable.trace_add("write", lambda *args: self.update_listvar())
            self.update_listvar()

    def update_listvar(self):
        try:
            values = ast.literal_eval(self.listvariable.get())
            if not isinstance(values, list):
                raise ValueError("Listvariable must contain a list.")
        except (SyntaxError, ValueError) as e:
            print(f"Error parsing listvariable: {e}")
            values = []

        self.delete_all()
        for option in values:
            self.insert("end", option=option)

    def select(self, index: Any):
        for button in self.buttons.values():
            button.configure(fg_color=self.button_fg_color)

        if isinstance(index, int):
            selected_button = self.buttons.get(index, None)
            if not selected_button:
                print(f"Index {index} out of range.")
                return
        else:
            selected_button = self.buttons.get(index, None)
            if not selected_button:
                print(f"No button found for index '{index}'.")
                return

        if self.multiple:
            if selected_button in self.selections:
                self.selections.remove(selected_button)
                selected_button.configure(fg_color=self.button_fg_color, hover=False)

                self.after(
                    100,
                    lambda b=selected_button: (
                        b.configure(hover=self.hover)
                        if str(b) in self.winfo_children()
                        else None
                    ),
                )
            else:
                self.selections.append(selected_button)
            for btn in self.selections:
                btn.configure(fg_color=self.select_color, hover=False)
                self.after(
                    100,
                    lambda b=btn: (
                        b.configure(hover=self.hover)
                        if str(b) in self.winfo_children()
                        else None
                    ),
                )
        else:
            self.selected = selected_button
            selected_button.configure(fg_color=self.select_color, hover=False)
            self.after(
                100,
                lambda b=selected_button: (
                    b.configure(hover=self.hover)
                    if str(b) in self.winfo_children()
                    else None
                ),
            )
            self.selected_index = list(self.buttons.values()).index(selected_button)

        if self.command:
            self.command(self.get())

        self.event_generate("<<ListboxSelect>>")

    def activate(self, index: Any):
        if isinstance(index, str) and index.lower() == "all":
            if self.multiple:
                for key in self.buttons.keys():
                    self.select(key)
            return

        if isinstance(index, str) and index.lower() == "end":
            index = -1

        if isinstance(index, int):
            selected_key = (
                list(self.buttons.keys())[index]
                if index >= 0
                else list(self.buttons.keys())[-1]
            )
            self.select(selected_key)

    def curselection(self) -> Tuple[int, ...]:
        if self.multiple:
            return tuple(
                idx for idx, btn in self.buttons.items() if btn in self.selections
            )
        else:
            if self.selected:
                return (list(self.buttons.values()).index(self.selected),)
            return ()

    def bind_selection(self, func: Any, add: str = "+"):
        self.bind("<Button-1>", lambda e: func(e), add=add)

    def unbind_selection(self):
        self.unbind("<Button-1>")

    def deselect(self, index: Any):
        if not self.multiple:
            if self.selected:
                self.selected.configure(fg_color=self.button_fg_color)
                self.selected = None
        else:
            if isinstance(index, int):
                button = self.buttons.get(index, None)
            else:
                button = self.buttons.get(index, None)
            if button and button in self.selections:
                self.selections.remove(button)
                button.configure(fg_color=self.button_fg_color)

    def deactivate(self, index: Any):
        if isinstance(index, str) and index.lower() == "all":
            if self.multiple:
                for key in list(self.buttons.keys()):
                    self.deselect(key)
            else:
                if self.selected:
                    self.deselect(self.selected_index)
        else:
            if isinstance(index, int):
                key = list(self.buttons.keys())[index]
                self.deselect(key)

    def insert(self, index: Any, option: str, update: bool = True, **kwargs):
        if isinstance(index, str) and index.lower() == "end":
            key = self.end_num
            self.end_num += 1
        elif isinstance(index, int):
            key = index
        else:
            key = index

        if key in self.buttons:
            self.buttons[key].destroy()

        button = ctk.CTkButton(
            self,
            text=option,
            fg_color=self.button_fg_color,
            anchor=self.justify,
            text_color=self.text_color,
            font=self.font,
            hover_color=self.hover_color,
            **kwargs,
        )
        button.configure(command=lambda k=key: self.select(k))
        button.grid(
            padx=0,
            pady=(0, 5),
            sticky="nsew",
            column=0,
            row=key if isinstance(key, int) else "n",
        )

        if self.multiple:
            button.bind("<Shift-Button-1>", lambda e, b=button: self.select_multiple(b))

        self.buttons[key] = button

        if update:
            self.update_idletasks()

        return button

    def select_multiple(self, button: ctk.CTkButton):
        selections = list(self.buttons.values())
        if not self.selections:
            self.select(self.get_key(button))
            return

        last = selections.index(self.selections[-1])
        current = selections.index(button)

        if last < current:
            for i in range(last + 1, current + 1):
                btn = selections[i]
                if btn not in self.selections:
                    self.select(self.get_key(btn))
        else:
            for i in range(current, last + 1):
                btn = selections[i]
                if btn not in self.selections:
                    self.select(self.get_key(btn))

    def get_key(self, button: ctk.CTkButton) -> Optional[int]:
        for key, btn in self.buttons.items():
            if btn == button:
                return key
        return None

    def destroy_all(self):
        for button in self.buttons.values():
            button.destroy()
        self._scrollbar.destroy()
        super().destroy()

    def delete_all(self):
        for key in list(self.buttons.keys()):
            self.buttons[key].destroy()
            del self.buttons[key]
        self.end_num = 0
        self.selections.clear()
        self.selected = None
        self.selected_index = 0

    def size(self) -> int:
        return len(self.buttons)

    def get(self, index: Optional[int] = None) -> Optional[Any]:
        if index is not None:
            if isinstance(index, str) and index.lower() == "all":
                return [btn.cget("text") for btn in self.buttons.values()]
            else:
                key = list(self.buttons.keys())[index]
                return self.buttons[key].cget("text")
        else:
            if self.multiple:
                return (
                    [btn.cget("text") for btn in self.selections]
                    if self.selections
                    else None
                )
            else:
                return self.selected.cget("text") if self.selected else None

    def configure_listbox(self, **kwargs):
        for key, value in kwargs.items():
            if key == "hover_color":
                self.hover_color = value
                for btn in self.buttons.values():
                    btn.configure(hover_color=self.hover_color)
            elif key == "button_color":
                self.button_fg_color = value
                for btn in self.buttons.values():
                    btn.configure(fg_color=self.button_fg_color)
            elif key == "highlight_color":
                self.select_color = value
                if self.selected:
                    self.selected.configure(fg_color=self.select_color)
                for btn in self.selections:
                    btn.configure(fg_color=self.select_color)
            elif key == "text_color":
                self.text_color = value
                for btn in self.buttons.values():
                    btn.configure(text_color=self.text_color)
            elif key == "font":
                self.font = value
                for btn in self.buttons.values():
                    btn.configure(font=self.font)
            elif key == "command":
                self.command = value
            elif key == "hover":
                self.hover = value
                for btn in self.buttons.values():
                    btn.configure(hover=self.hover)
            elif key == "justify":
                justify_map = {"left": "w", "right": "e", "center": "c"}
                self.justify = justify_map.get(value.lower(), "w")
                for btn in self.buttons.values():
                    btn.configure(anchor=self.justify)
            elif key == "height":
                self._scrollbar.configure(height=value)
            elif key == "multiple_selection":
                self.multiple = value
            elif key == "options":
                self.delete_all()
                for option in value:
                    self.insert("end", option=option)
        self.update_idletasks()

    def cget_listbox(self, param: str) -> Any:
        config_map = {
            "hover_color": self.hover_color,
            "button_color": self.button_fg_color,
            "highlight_color": self.select_color,
            "text_color": self.text_color,
            "font": self.font,
            "hover": self.hover,
            "justify": self.justify,
        }
        return config_map.get(param, super().cget(param))


class Condition:
    def __init__(self, condition_type: str, value: str):
        self.condition_type = condition_type
        self.value = value

    def is_met(self) -> bool:
        print(f"Checking condition: {self.condition_type} with value: {self.value}")
        if self.condition_type == "window_active":
            active_window = gw.getActiveWindow()
            if active_window:
                result = self.value.lower() in active_window.title.lower()
                print(
                    f"Active window '{active_window.title}' contains '{self.value}': {result}"
                )
                return result
            else:
                print("No active window found.")
                return False
        return True


class Action:
    def __init__(
        self,
        action_type: str,
        value: str,
        interval: Optional[float] = None,
        min_interval: Optional[float] = None,
        max_interval: Optional[float] = None,
        target_window: Optional[int] = None,
    ):
        self.action_type = action_type
        self.value = value
        self.interval = interval
        self.min_interval = min_interval
        self.max_interval = max_interval
        self.running = True
        self.target_window = target_window

    def perform(self, global_condition: Optional[Condition] = None):
        print(f"Performing action: {self.action_type} - {self.value}")

        if self.target_window and not self.is_window_valid(self.target_window):
            print(f"Target window {self.target_window} is not valid or is minimized.")
            self.running = False

            return

        if global_condition and not global_condition.is_met():
            print(
                f"Global condition '{global_condition.condition_type}' not met. Waiting..."
            )
            while self.running and not global_condition.is_met():
                time.sleep(0.5)
                print(
                    f"Checking global condition '{global_condition.condition_type}' again..."
                )
        if self.action_type == "key_press":
            self.press_key(self.value)
        elif self.action_type == "mouse_click":
            self.mouse_click(self.value)

    def press_key(self, key_name: str):
        try:
            if not self.target_window:

                mapped_key = KEY_NAME_MAPPING.get(key_name.lower(), key_name)
                print(f"Pressing key: {mapped_key}")
                keyboard_controller.press(mapped_key)
                time.sleep(0.1)
                keyboard_controller.release(mapped_key)
                print(f"Key '{mapped_key}' pressed successfully.")
            else:
                hwnd = self.target_window
                print(f"Target hwnd: {hwnd}")
                if not isinstance(hwnd, int):
                    hwnd = int(hwnd)
                    print(f"Handle converted to int: {hwnd}")

                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

                vk_code = win32api.VkKeyScan(key_name.upper()) & 0xFF
                scan_code = win32api.MapVirtualKey(vk_code, 0)

                lParam_down = 1 | (scan_code << 16)

                extended_keys = [
                    "right",
                    "left",
                    "insert",
                    "delete",
                    "home",
                    "end",
                    "pageup",
                    "pagedown",
                    "up",
                    "down",
                    "left",
                    "right",
                ]
                if key_name.lower() in extended_keys:
                    lParam_down |= 1 << 24

                lParam_up = 1 | (scan_code << 16) | (1 << 30) | (1 << 31)
                if key_name.lower() in extended_keys:
                    lParam_up |= 1 << 24

                print(f"Sending WM_KEYDOWN for '{key_name}' with lParam: {lParam_down}")
                win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, vk_code, lParam_down)
                print(f"Sending WM_KEYUP for '{key_name}' with lParam: {lParam_up}")
                win32gui.SendMessage(hwnd, win32con.WM_KEYUP, vk_code, lParam_up)
                print(f"Key '{key_name}' sent to target window via SendMessage.")
        except Exception as e:
            print(f"Error pressing key '{key_name}': {e}")

    def mouse_click(self, button: str):
        try:
            if not self.target_window:

                pyautogui.click(button=button)
                print(f"Clicked mouse button: {button}")
            else:
                hwnd = self.target_window
                print(f"Target hwnd: {hwnd}")
                if not isinstance(hwnd, int):
                    hwnd = int(hwnd)
                    print(f"Handle converted to int: {hwnd}")

                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

                left, top, right, bottom = win32gui.GetClientRect(hwnd)
                width = right - left
                height = bottom - top

                x = width // 2
                y = height // 2
                lParam = win32api.MAKELONG(x, y)

                if button.lower() == "left":
                    msg_down = win32con.WM_LBUTTONDOWN
                    msg_up = win32con.WM_LBUTTONUP
                    wParam = win32con.MK_LBUTTON
                elif button.lower() == "right":
                    msg_down = win32con.WM_RBUTTONDOWN
                    msg_up = win32con.WM_RBUTTONUP
                    wParam = win32con.MK_RBUTTON
                elif button.lower() == "middle":
                    msg_down = win32con.WM_MBUTTONDOWN
                    msg_up = win32con.WM_MBUTTONUP
                    wParam = win32con.MK_MBUTTON
                else:
                    print(f"Unknown mouse button: {button}")
                    return

                print(f"Sending {msg_down} for '{button}' with lParam: {lParam}")
                win32gui.SendMessage(hwnd, msg_down, wParam, lParam)
                print(f"Sending {msg_up} for '{button}' with lParam: {lParam}")
                win32gui.SendMessage(hwnd, msg_up, 0, lParam)
                print(f"Mouse '{button}' click sent to target window via SendMessage.")
        except Exception as e:
            print(f"Error clicking mouse button '{button}': {e}")

    @staticmethod
    def is_window_valid(hwnd: int) -> bool:
        return win32gui.IsWindow(hwnd) and not win32gui.IsIconic(hwnd)

    def get_interval(self) -> Optional[float]:
        if self.min_interval is not None and self.max_interval is not None:
            return random.uniform(self.min_interval, self.max_interval)
        else:
            return self.interval


class CustomIntervalInputDialog(ctk.CTkToplevel):
    def __init__(
        self, parent: ctk.CTk, title: str = "Set Interval", randomize: bool = False
    ):
        super().__init__(parent)
        self.title(title)
        self.geometry("350x350" if randomize else "350x150")
        self.resizable(False, False)
        self.grab_set()
        self.value: Optional[
            Tuple[Optional[float], Optional[float], Optional[float]]
        ] = None
        self.randomize = randomize

        if not self.randomize:
            self.interval_label = ctk.CTkLabel(
                self, text="Interval after action (seconds):"
            )
            self.interval_label.pack(pady=10, padx=10)
            self.interval_entry = ctk.CTkEntry(self)
            self.interval_entry.pack(pady=5, padx=10)
        else:
            self.min_interval_label = ctk.CTkLabel(
                self, text="Minimum interval (seconds):"
            )
            self.min_interval_label.pack(pady=10, padx=10)
            self.min_interval_entry = ctk.CTkEntry(self)
            self.min_interval_entry.pack(pady=5, padx=10)

            self.max_interval_label = ctk.CTkLabel(
                self, text="Maximum interval (seconds):"
            )
            self.max_interval_label.pack(pady=10, padx=10)
            self.max_interval_entry = ctk.CTkEntry(self)
            self.max_interval_entry.pack(pady=5, padx=10)

        self.button_frame = ctk.CTkFrame(self)
        self.button_frame.pack(pady=10)

        self.ok_button = ctk.CTkButton(self.button_frame, text="OK", command=self.on_ok)
        self.ok_button.grid(row=0, column=0, padx=10)

        self.cancel_button = ctk.CTkButton(
            self.button_frame, text="Cancel", command=self.on_cancel
        )
        self.cancel_button.grid(row=0, column=1, padx=10)

    def on_ok(self):
        try:
            if not self.randomize:
                interval = float(self.interval_entry.get())
                if interval <= 0:
                    raise ValueError("Interval must be positive.")
                self.value = (interval, None, None)
            else:
                min_interval = float(self.min_interval_entry.get())
                max_interval = float(self.max_interval_entry.get())
                if min_interval <= 0 or max_interval <= 0:
                    raise ValueError("Intervals must be positive numbers.")
                if min_interval >= max_interval:
                    raise ValueError(
                        "Minimum interval must be less than maximum interval."
                    )
                self.value = (None, min_interval, max_interval)
            self.grab_release()
            self.destroy()
        except ValueError as e:
            self.show_error(str(e))

    def on_cancel(self):
        self.value = None
        self.grab_release()
        self.destroy()

    def show_error(self, message: str):
        error_dialog = CustomMessageBox(self, title="Error", message=message)
        error_dialog.wait_window()


class CustomMessageBox(ctk.CTkToplevel):
    def __init__(
        self,
        parent: ctk.CTk,
        title: str = "Message",
        message: str = "",
        confirm: bool = False,
    ):
        super().__init__(parent)
        self.title(title)
        self.geometry("350x200")
        self.resizable(False, False)
        self.grab_set()

        self.label = ctk.CTkLabel(self, text=message)
        self.label.pack(pady=20, padx=20)

        self.confirmed = False

        self.button_frame = ctk.CTkFrame(self)
        self.button_frame.pack(pady=10)

        if confirm:
            self.yes_button = ctk.CTkButton(
                self.button_frame, text="Yes", command=self.on_yes
            )
            self.yes_button.grid(row=0, column=0, padx=10)

            self.no_button = ctk.CTkButton(
                self.button_frame, text="No", command=self.on_no
            )
            self.no_button.grid(row=0, column=1, padx=10)
        else:
            self.ok_button = ctk.CTkButton(
                self.button_frame, text="OK", command=self.on_ok
            )
            self.ok_button.pack(pady=10)

    def on_ok(self):
        self.grab_release()
        self.destroy()

    def on_yes(self):
        self.confirmed = True
        self.grab_release()
        self.destroy()

    def on_no(self):
        self.confirmed = False
        self.grab_release()
        self.destroy()


class ConditionDialog(ctk.CTkToplevel):
    def __init__(self, parent: ctk.CTk):
        super().__init__(parent)
        self.title("Set Condition")
        self.geometry("350x300")
        self.resizable(False, False)
        self.grab_set()
        self.condition: Optional[Condition] = None

        self.condition_type_var = ctk.StringVar(value="None")

        self.condition_type_label = ctk.CTkLabel(self, text="Condition Type:")
        self.condition_type_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.condition_type_option = ctk.CTkOptionMenu(
            self,
            variable=self.condition_type_var,
            values=["None", "Window Active"],
            command=self.update_value_widget,
        )
        self.condition_type_option.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        self.value_label = ctk.CTkLabel(self, text="Value:")
        self.value_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        self.value_widget = ctk.CTkLabel(self, text="")
        self.value_widget.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        self.button_frame = ctk.CTkFrame(self)
        self.button_frame.grid(row=2, column=0, columnspan=2, pady=10)

        self.ok_button = ctk.CTkButton(self.button_frame, text="OK", command=self.on_ok)
        self.ok_button.grid(row=0, column=0, padx=10)

        self.cancel_button = ctk.CTkButton(
            self.button_frame, text="Cancel", command=self.on_cancel
        )
        self.cancel_button.grid(row=0, column=1, padx=10)

        self.grid_columnconfigure(1, weight=1)

    def update_value_widget(self, *args):

        self.value_widget.grid_forget()

        condition_type = self.condition_type_var.get()
        if condition_type == "Window Active":

            window_titles = [title for title in gw.getAllTitles() if title.strip()]
            if not window_titles:
                window_titles = ["No windows found"]
            self.value_var = ctk.StringVar(value=window_titles[0])
            self.value_widget = ctk.CTkOptionMenu(
                self, variable=self.value_var, values=window_titles
            )
        else:
            self.value_var = ctk.StringVar(value="")
            self.value_widget = ctk.CTkLabel(self, text="No value required")
        self.value_widget.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

    def on_ok(self):
        condition_type = self.condition_type_var.get()
        if condition_type != "None":
            value = self.value_var.get()
            self.condition = Condition(condition_type.lower().replace(" ", "_"), value)
            print(
                f"Condition set: {self.condition.condition_type} - {self.condition.value}"
            )
        self.grab_release()
        self.destroy()

    def on_cancel(self):
        self.grab_release()
        self.destroy()


class AutoAction:
    def __init__(self, root: ctk.CTk):
        self.root = root
        self.root.title("ActionFlow by Crashdaemon")
        self.root.geometry("650x1300")
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("dark-blue")

        self.activation_key: Optional[str] = None
        self.action_sequence: List[Action] = []
        self.start_delay: float = 0.0
        self.running: bool = False
        self.action_thread: Optional[threading.Thread] = None
        self.repeat_option = ctk.StringVar(value="infinite")
        self.loading_preset: bool = False
        self.recording: bool = False
        self.recorded_events: List[Tuple[str, str, float]] = []
        self.global_condition: Optional[Condition] = None
        self.last_event_time = None

        self.last_event_time: Optional[float] = None
        self.keyboard_listener = None
        self.mouse_listener = None
        self.window_handles = []
        self.target_window = None

        self.create_widgets()
        self.load_last_preset_on_startup()

    def create_widgets(self):

        self.profile_frame = ctk.CTkFrame(self.root)
        self.profile_frame.pack(pady=10)

        self.new_profile_button = ctk.CTkButton(
            self.profile_frame, text="New Profile", command=self.new_profile
        )
        self.new_profile_button.grid(row=0, column=0, padx=10)

        self.save_profile_button = ctk.CTkButton(
            self.profile_frame, text="Save Profile", command=self.save_preset
        )
        self.save_profile_button.grid(row=0, column=1, padx=10)

        self.load_profile_button = ctk.CTkButton(
            self.profile_frame, text="Load Profile", command=self.load_preset
        )
        self.load_profile_button.grid(row=0, column=2, padx=10)

        self.control_frame = ctk.CTkFrame(self.root)
        self.control_frame.pack(pady=20)

        self.activation_key_button = ctk.CTkButton(
            self.control_frame,
            text="Set Activation Key",
            command=self.set_activation_key,
        )
        self.activation_key_button.grid(row=0, column=0, padx=10)

        self.start_button = ctk.CTkButton(
            self.control_frame, text="Start", command=self.start_action
        )
        self.start_button.grid(row=0, column=1, padx=10)

        self.stop_button = ctk.CTkButton(
            self.control_frame, text="Stop", command=self.stop_action
        )
        self.stop_button.grid(row=0, column=2, padx=10)

        self.activation_key_label = ctk.CTkLabel(
            self.control_frame, text="Activation Key: None"
        )
        self.activation_key_label.grid(row=1, column=0, columnspan=3, pady=5)

        self.action_sequence_label = ctk.CTkLabel(
            self.root, text="Action Sequence", font=ctk.CTkFont(size=16, weight="bold")
        )
        self.action_sequence_label.pack(pady=(20, 10))

        self.action_sequence_frame = ctk.CTkFrame(self.root)
        self.action_sequence_frame.pack(pady=5)

        self.add_key_press_button = ctk.CTkButton(
            self.action_sequence_frame,
            text="Add Key Press",
            command=self.add_key_press_action,
        )
        self.add_key_press_button.grid(row=0, column=0, padx=10, pady=5)

        self.add_mouse_click_button = ctk.CTkButton(
            self.action_sequence_frame,
            text="Add Mouse Click",
            command=self.add_mouse_click_action,
        )
        self.add_mouse_click_button.grid(row=0, column=1, padx=10, pady=5)

        self.delete_sequence_button = ctk.CTkButton(
            self.action_sequence_frame,
            text="Delete Action Sequence",
            command=self.delete_action_sequence,
        )
        self.delete_sequence_button.grid(row=0, column=2, padx=10, pady=5)

        self.edit_action_button = ctk.CTkButton(
            self.action_sequence_frame,
            text="Edit Selected Action",
            command=self.edit_selected_action,
        )
        self.edit_action_button.grid(row=0, column=3, padx=10, pady=5)

        self.action_listbox = CTkListbox(
            self.root, height=200, width=600, multiple_selection=True
        )
        self.action_listbox.pack(pady=5)

        self.record_frame = ctk.CTkFrame(self.root)
        self.record_frame.pack(pady=10)

        self.start_record_button = ctk.CTkButton(
            self.record_frame, text="Start Recording", command=self.start_recording
        )
        self.start_record_button.grid(row=0, column=0, padx=10)

        self.stop_record_button = ctk.CTkButton(
            self.record_frame,
            text="Stop Recording",
            command=self.stop_recording,
            state=ctk.DISABLED,
        )
        self.stop_record_button.grid(row=0, column=1, padx=10)

        self.global_condition_label = ctk.CTkLabel(
            self.root, text="Global Condition", font=ctk.CTkFont(size=16, weight="bold")
        )
        self.global_condition_label.pack(pady=(20, 10))

        self.global_condition_frame = ctk.CTkFrame(self.root)
        self.global_condition_frame.pack(pady=5)

        self.set_global_condition_button = ctk.CTkButton(
            self.global_condition_frame,
            text="Set Global Condition",
            command=self.set_global_condition,
        )
        self.set_global_condition_button.pack(pady=5)

        self.current_global_condition_label = ctk.CTkLabel(
            self.global_condition_frame, text="No global condition set."
        )
        self.current_global_condition_label.pack(pady=5)

        self.window_selection_label = ctk.CTkLabel(
            self.root, text="Target Window", font=ctk.CTkFont(size=16, weight="bold")
        )
        self.window_selection_label.pack(pady=(20, 10))

        self.window_selection_frame = ctk.CTkFrame(self.root)
        self.window_selection_frame.pack(pady=5)

        self.refresh_windows_button = ctk.CTkButton(
            self.window_selection_frame,
            text="Refresh Windows",
            command=self.refresh_window_list,
        )
        self.refresh_windows_button.grid(row=0, column=0, padx=10)

        self.selected_window_var = ctk.StringVar(value="Select a window")
        self.window_option_menu = ctk.CTkOptionMenu(
            self.window_selection_frame,
            variable=self.selected_window_var,
            values=[],
            command=lambda _: self.set_target_window(),
        )
        self.window_option_menu.grid(row=0, column=1, padx=10)

        self.target_window_option_label = ctk.CTkLabel(
            self.window_selection_frame, text="Enable Targeted Window Actions:"
        )
        self.target_window_option_label.grid(
            row=1, column=0, padx=10, pady=10, sticky="w"
        )

        self.use_target_window_var = ctk.BooleanVar(value=False)
        self.use_target_window_switch = ctk.CTkSwitch(
            self.window_selection_frame,
            text="Use Targeted Window",
            variable=self.use_target_window_var,
            command=self.toggle_target_window_options,
        )
        self.use_target_window_switch.grid(
            row=1, column=1, padx=10, pady=10, sticky="w"
        )

        self.refresh_windows_button.configure(state=ctk.DISABLED)
        self.window_option_menu.configure(state="disabled")

        self.repeat_label = ctk.CTkLabel(self.root, text="Repeat Sequence:")
        self.repeat_label.pack(pady=5)

        self.repeat_frame = ctk.CTkFrame(self.root)
        self.repeat_frame.pack(pady=5)

        self.repeat_infinite_radio = ctk.CTkRadioButton(
            self.repeat_frame,
            text="Infinite",
            variable=self.repeat_option,
            value="infinite",
            command=self.toggle_repeat_count_entry,
        )
        self.repeat_infinite_radio.grid(row=0, column=0, padx=10, pady=5)

        self.repeat_custom_radio = ctk.CTkRadioButton(
            self.repeat_frame,
            text="Custom",
            variable=self.repeat_option,
            value="custom",
            command=self.toggle_repeat_count_entry,
        )
        self.repeat_custom_radio.grid(row=0, column=1, padx=10, pady=5)

        self.repeat_count_entry = ctk.CTkEntry(self.repeat_frame)
        self.repeat_count_entry.insert(0, "1")
        self.repeat_count_entry.grid(row=1, column=0, columnspan=2, pady=5)
        self.repeat_count_entry.grid_remove()

        self.delay_label = ctk.CTkLabel(
            self.root, text="Delay before starting (seconds):"
        )
        self.delay_label.pack(pady=(20, 5))

        self.delay_entry = ctk.CTkEntry(self.root)
        self.delay_entry.insert(0, "0.0")
        self.delay_entry.pack(pady=5)

        self.status_label = ctk.CTkLabel(
            self.root, text="Status: Stopped", text_color="red"
        )
        self.status_label.pack(pady=10)

        self.log_label = ctk.CTkLabel(self.root, text="Action Log:")
        self.log_label.pack(pady=5)

        self.log_textbox = ctk.CTkTextbox(self.root, width=600, height=150)
        self.log_textbox.pack(pady=5)
        self.log_textbox.configure(state="disabled")

        self.clear_log_button = ctk.CTkButton(
            self.root, text="Clear Log", command=self.clear_log
        )
        self.clear_log_button.pack(pady=5)

        self.appearance_label = ctk.CTkLabel(
            self.root,
            text="Appearance Settings",
            font=ctk.CTkFont(size=16, weight="bold"),
        )
        self.appearance_label.pack(pady=(20, 10))

        self.appearance_frame = ctk.CTkFrame(self.root)
        self.appearance_frame.pack(pady=5)

        self.mode_label = ctk.CTkLabel(self.appearance_frame, text="Mode:")
        self.mode_label.grid(row=0, column=0, padx=5, pady=5)

        self.appearance_mode_option = ctk.CTkOptionMenu(
            self.appearance_frame,
            values=["System", "Light", "Dark"],
            command=self.change_appearance_mode,
        )
        self.appearance_mode_option.set("System")
        self.appearance_mode_option.grid(row=0, column=1, padx=5, pady=5)

    def load_last_preset_on_startup(self):
        try:
            with open("config.json", "r") as config_file:
                config = json.load(config_file)
                last_preset_path = config.get("last_preset", None)
                if last_preset_path:
                    self.loading_preset = True
                    self.load_preset_from_path(last_preset_path)
        except FileNotFoundError:
            pass
        except Exception as e:
            self.show_message("Error", f"Failed to load last preset: {e}")
        finally:
            self.loading_preset = False

    def load_preset_from_path(self, file_path: str):
        try:
            with open(file_path, "r") as f:
                preset = json.load(f)
            self.use_target_window_var.set(preset.get("use_target_window", False))
            self.toggle_target_window_options()

            self.activation_key = preset.get("activation_key", None)
            self.activation_key_label.configure(
                text=(
                    f"Activation Key: {self.activation_key}"
                    if self.activation_key
                    else "Activation Key: None"
                )
            )
            if self.activation_key:
                keyboard.add_hotkey(self.activation_key, self.toggle_action)

            self.action_sequence = []
            for action_dict in preset.get("action_sequence", []):
                target_window = None
                if action_dict.get("target_window_title"):
                    try:
                        windows = gw.getWindowsWithTitle(
                            action_dict["target_window_title"]
                        )
                        if windows:
                            target_window = windows[0]._hWnd
                        else:
                            print(
                                f"Target window '{action_dict['target_window_title']}' not found."
                            )
                    except Exception as e:
                        print(
                            f"Failed to find target window '{action_dict['target_window_title']}': {e}"
                        )
                action = Action(
                    action_type=action_dict["action_type"],
                    value=action_dict["value"],
                    interval=action_dict.get("interval", None),
                    min_interval=action_dict.get("min_interval", None),
                    max_interval=action_dict.get("max_interval", None),
                    target_window=target_window,
                )
                self.action_sequence.append(action)
            self.update_action_sequence_display()
            self.delay_entry.delete(0, ctk.END)
            self.delay_entry.insert(0, preset.get("start_delay", "0.0"))

            self.repeat_option.set(preset.get("repeat_option", "infinite"))
            if self.repeat_option.get() == "custom":
                self.repeat_count_entry.delete(0, ctk.END)
                self.repeat_count_entry.insert(0, preset.get("repeat_count", "1"))
            self.global_condition = None
            if preset.get("global_condition", None):
                cond = preset["global_condition"]
                self.global_condition = Condition(cond["condition_type"], cond["value"])
                self.current_global_condition_label.configure(
                    text=f"Condition: {self.global_condition.condition_type.capitalize()} ({self.global_condition.value})"
                )
            else:
                self.current_global_condition_label.configure(
                    text="No global condition set."
                )

            if self.use_target_window_var.get() and preset.get("target_window_title"):
                try:
                    windows = gw.getWindowsWithTitle(preset["target_window_title"])
                    if windows:
                        window = windows[0]
                        self.target_window = window._hWnd

                        window_list = [
                            f"{idx + 1}. {window.title}"
                            for idx, window in enumerate(gw.getAllWindows())
                            if window.title
                        ]
                        target_title = preset["target_window_title"]

                        target_index = next(
                            (
                                idx
                                for idx, title in enumerate(
                                    [w.split(". ", 1)[1] for w in window_list]
                                )
                                if title == target_title
                            ),
                            None,
                        )
                        if target_index is not None:
                            self.selected_window_var.set(window_list[target_index])
                        else:
                            self.selected_window_var.set("Select a window")
                    else:
                        self.target_window = None
                        self.selected_window_var.set("Select a window")
                        print(
                            f"Target window '{preset['target_window_title']}' not found."
                        )
                except Exception as e:
                    self.show_message(
                        "Error", f"Failed to reconnect to target window: {e}"
                    )
                    print(f"Failed to reconnect to target window: {e}")
                    self.target_window = None
                    self.selected_window_var.set("Select a window")

            self.toggle_repeat_count_entry()
            print(f"Preset loaded from {file_path}")
        except Exception as e:
            self.show_message("Error", f"Failed to load preset: {e}")
            print(f"Failed to load preset: {e}")

    def set_global_condition(self):
        condition_dialog = ConditionDialog(self.root)
        self.root.wait_window(condition_dialog)
        condition = condition_dialog.condition
        if condition:
            self.global_condition = condition
            self.current_global_condition_label.configure(
                text=f"Condition: {condition.condition_type.capitalize()} ({condition.value})"
            )
            print(
                f"Global condition set: {condition.condition_type} - {condition.value}"
            )
        else:
            self.global_condition = None
            self.current_global_condition_label.configure(
                text="No global condition set."
            )
            print("Global condition cleared.")

    def toggle_repeat_count_entry(self):
        if self.repeat_option.get() == "custom":
            self.repeat_count_entry.grid()
        else:
            self.repeat_count_entry.grid_remove()

    def set_activation_key(self):
        self.activation_key_label.configure(text="Press a key for Activation...")
        self.activation_key_button.configure(state=ctk.DISABLED)
        self.add_key_press_button.configure(state=ctk.DISABLED)
        self.add_mouse_click_button.configure(state=ctk.DISABLED)
        self.start_record_button.configure(state=ctk.DISABLED)

        def on_key(event):
            self.activation_key = event.name
            self.activation_key_label.configure(
                text=f"Activation Key: {self.activation_key}"
            )
            keyboard.unhook_all_hotkeys()
            keyboard.add_hotkey(self.activation_key, self.toggle_action)
            keyboard.unhook(on_key)
            self.activation_key_button.configure(state=ctk.NORMAL)
            self.add_key_press_button.configure(state=ctk.NORMAL)
            self.add_mouse_click_button.configure(state=ctk.NORMAL)
            self.start_record_button.configure(state=ctk.NORMAL)
            print(f"Activation key set to: {self.activation_key}")
            return False

        keyboard.hook(on_key)

    def add_key_press_action(self):
        self.add_key_press_button.configure(state=ctk.DISABLED)
        self.add_mouse_click_button.configure(state=ctk.DISABLED)
        self.activation_key_button.configure(state=ctk.DISABLED)
        self.start_record_button.configure(state=ctk.DISABLED)

        def on_key(event):
            key_name = event.name
            keyboard.unhook(on_key)
            self.root.after(0, self.add_action_to_sequence, key_name, "key_press")
            print(f"Key press action added: {key_name}")
            return False

        keyboard.hook(on_key)

    def add_mouse_click_action(self):
        self.add_key_press_button.configure(state=ctk.DISABLED)
        self.add_mouse_click_button.configure(state=ctk.DISABLED)
        self.activation_key_button.configure(state=ctk.DISABLED)
        self.start_record_button.configure(state=ctk.DISABLED)
        self.select_mouse_button()

    def select_mouse_button(self):
        button_window = ctk.CTkToplevel(self.root)
        button_window.title("Select Mouse Button")
        button_window.geometry("250x200")
        button_window.resizable(False, False)
        button_window.grab_set()

        label = ctk.CTkLabel(button_window, text="Select Mouse Button:")
        label.pack(pady=10)

        mouse_buttons = ["left", "right", "middle"]
        for btn in mouse_buttons:
            btn_widget = ctk.CTkButton(
                button_window,
                text=btn.capitalize(),
                command=lambda b=btn: self.on_mouse_button_selected(b, button_window),
            )
            btn_widget.pack(pady=5)

    def on_mouse_button_selected(self, button: str, window: ctk.CTkToplevel):
        window.destroy()
        self.add_action_to_sequence(button, "mouse_click")
        print(f"Mouse click action added: {button}")

    def add_action_to_sequence(self, value: str, action_type: str):
        try:

            if self.use_target_window_var.get():
                target_window = self.target_window
            else:
                target_window = None

            action = Action(action_type, value, target_window=target_window)
            self.action_sequence.append(action)
            self.update_action_sequence_display()
            print(f"Action added: {action_type} - {value}")

            last_index = len(self.action_sequence) - 1
            self.action_listbox.select(last_index)

            self.root.after(100, self.edit_selected_action)

        except Exception as e:
            self.show_message("Error", f"Failed to add action: {e}")
            print(f"Failed to add action: {e}")
        finally:
            self.add_key_press_button.configure(state=ctk.NORMAL)
            self.add_mouse_click_button.configure(state=ctk.NORMAL)
            self.activation_key_button.configure(state=ctk.NORMAL)
            self.start_record_button.configure(state=ctk.NORMAL)

    def update_action_sequence_display(self):
        self.action_listbox.delete_all()
        for idx, action in enumerate(self.action_sequence):

            action_text = f"{idx + 1}. {action.action_type.capitalize()}: {action.value} | Interval: "
            if action.interval is not None:
                action_text += f"{action.interval:.2f}s"
            elif action.min_interval is not None and action.max_interval is not None:
                action_text += (
                    f"Random ({action.min_interval:.2f}s - {action.max_interval:.2f}s)"
                )
            self.action_listbox.insert("end", option=action_text)

    def start_action(self):
        if self.running:
            self.show_message("Info", "Auto action is already running!")
            return
        try:
            self.start_delay = float(self.delay_entry.get())
            if self.start_delay < 0:
                raise ValueError("Delay cannot be negative.")
        except ValueError as e:
            self.show_message("Error", f"Invalid input: {e}")
            return
        if not self.action_sequence:
            self.show_message("Error", "Please add actions to the action sequence.")
            return

        self.running = True
        self.status_label.configure(text="Status: Starting...", text_color="orange")
        self.action_thread = threading.Thread(target=self.perform_action, daemon=True)
        self.action_thread.start()
        print("Action thread started.")

    def stop_action(self):
        self.running = False
        self.status_label.configure(text="Status: Stopped", text_color="red")
        print("Action thread stopped.")

    def toggle_action(self):
        if self.running:
            self.stop_action()
        else:
            self.start_action()

    def perform_action(self):
        for i in range(int(self.start_delay), 0, -1):
            if not self.running:
                break
            self.status_label.configure(
                text=f"Starting in {i} seconds...", text_color="orange"
            )
            print(f"Starting in {i} seconds...")
            time.sleep(1)

        if not self.running:
            self.status_label.configure(text="Status: Stopped", text_color="red")
            print("Action stopped before starting.")
            return

        self.status_label.configure(text="Status: Running", text_color="green")
        print("Action is now running.")

        repeat = self.repeat_option.get()
        repeat_count = 0
        max_repeats = None

        if repeat == "custom":
            try:
                max_repeats = int(self.repeat_count_entry.get())
                if max_repeats <= 0:
                    raise ValueError("Repeat count must be a positive integer.")
            except ValueError as e:
                self.show_message("Error", f"Invalid repeat count: {e}")
                self.stop_action()
                return

        while self.running:
            if repeat == "custom" and repeat_count >= max_repeats:
                print("Reached maximum repeat count.")
                break
            try:
                for action in self.action_sequence:
                    if not self.running:
                        break
                    action.perform(global_condition=self.global_condition)
                    current_time = time.strftime("%H:%M:%S")

                    if action.action_type == "key_press":
                        log_message = f"[{current_time}] Pressed key: {action.value}"
                    elif action.action_type == "mouse_click":
                        log_message = (
                            f"[{current_time}] Clicked mouse button: {action.value}"
                        )
                    else:
                        log_message = (
                            f"[{current_time}] Unknown action: {action.action_type}"
                        )

                    if action.interval is not None:
                        interval = action.interval
                        log_message += f" | Interval: {interval:.2f}s"
                    elif (
                        action.min_interval is not None
                        and action.max_interval is not None
                    ):
                        interval = action.get_interval()
                        log_message += f" | Interval: Random ({action.min_interval:.2f}s - {action.max_interval:.2f}s) (used {interval:.2f}s)"
                    else:
                        interval = 1.0

                    self.log_textbox.configure(state="normal")
                    self.log_textbox.insert(ctk.END, log_message + "\n")
                    self.log_textbox.see(ctk.END)
                    self.log_textbox.configure(state="disabled")
                    print(log_message)

                    time.sleep(interval)

                repeat_count += 1
                if repeat == "once":
                    print("Completed single iteration of actions.")
                    break
            except Exception as e:
                self.show_message("Error", f"An error occurred: {e}")
                break

        self.stop_action()

    def save_preset(self):
        preset = {
            "use_target_window": self.use_target_window_var.get(),
            "activation_key": self.activation_key,
            "action_sequence": [
                {
                    "action_type": action.action_type,
                    "value": action.value,
                    "interval": action.interval,
                    "min_interval": action.min_interval,
                    "max_interval": action.max_interval,
                    "target_window_title": (
                        win32gui.GetWindowText(action.target_window)
                        if action.target_window
                        else None
                    ),
                }
                for action in self.action_sequence
            ],
            "start_delay": self.start_delay,
            "repeat_option": self.repeat_option.get(),
            "repeat_count": (
                self.repeat_count_entry.get()
                if self.repeat_option.get() == "custom"
                else None
            ),
            "global_condition": (
                {
                    "condition_type": self.global_condition.condition_type,
                    "value": self.global_condition.value,
                }
                if self.global_condition
                else None
            ),
            "target_window_title": (
                win32gui.GetWindowText(self.target_window)
                if self.target_window
                else None
            ),
        }
        file_path = ctk.filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            title="Save Preset",
        )
        if file_path:
            try:
                with open(file_path, "w") as f:
                    json.dump(preset, f, indent=4)
                self.save_last_preset_path(file_path)
                self.show_message(
                    "Preset Saved", "Your preset has been saved successfully."
                )
                print(f"Preset saved to {file_path}")
            except Exception as e:
                self.show_message("Error", f"Failed to save preset: {e}")
                print(f"Failed to save preset: {e}")

    def save_last_preset_path(self, file_path: str):
        config = {"last_preset": file_path}
        try:
            with open("config.json", "w") as config_file:
                json.dump(config, config_file)
            print(f"Last preset path saved: {file_path}")
        except Exception as e:
            print(f"Failed to save last preset path: {e}")

    def load_preset(self):
        file_path = ctk.filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json")], title="Load Preset"
        )
        if file_path:
            try:
                self.loading_preset = True
                self.load_preset_from_path(file_path)
                print(f"Preset loaded from {file_path}")
            finally:
                self.loading_preset = False

    def new_profile(self):
        self.activation_key = None
        self.activation_key_label.configure(text="Activation Key: None")
        self.action_sequence = []
        self.update_action_sequence_display()
        self.delay_entry.delete(0, ctk.END)
        self.delay_entry.insert(0, "0.0")

        self.repeat_option.set("infinite")
        self.repeat_count_entry.delete(0, ctk.END)
        self.repeat_count_entry.insert(0, "1")

        self.toggle_repeat_count_entry()
        self.global_condition = None
        self.current_global_condition_label.configure(text="No global condition set.")
        print("New profile created.")

    def show_confirmation(self, title: str, message: str) -> bool:
        confirm_dialog = CustomMessageBox(
            self.root, title=title, message=message, confirm=True
        )
        self.root.wait_window(confirm_dialog)
        return getattr(confirm_dialog, "confirmed", False)

    def edit_selected_action(self):
        selected_indices = self.action_listbox.curselection()
        if not selected_indices:
            self.show_message("Info", "Please select an action to edit.")
            return
        index = selected_indices[0]
        action = self.action_sequence[index]

        edit_window = ctk.CTkToplevel(self.root)
        edit_window.title("Edit Action")
        edit_window.geometry("500x300")
        edit_window.resizable(False, False)
        edit_window.grab_set()

        action_type_label = ctk.CTkLabel(edit_window, text="Action Type:")
        action_type_label.grid(row=0, column=0, padx=10, pady=(10, 0), sticky="w")

        action_type_var = ctk.StringVar(value=action.action_type)
        action_type_option = ctk.CTkOptionMenu(
            edit_window,
            variable=action_type_var,
            values=["key_press", "mouse_click"],
            command=lambda _: self.update_edit_value_widget(
                action_type_var, value_var, edit_window
            ),
        )
        action_type_option.grid(row=0, column=1, padx=10, pady=(10, 0), sticky="ew")

        value_label = ctk.CTkLabel(edit_window, text="Value:")
        value_label.grid(row=1, column=0, padx=10, pady=(10, 0), sticky="w")

        value_var = ctk.StringVar(value=action.value)
        value_entry = ctk.CTkEntry(edit_window, textvariable=value_var)
        value_entry.grid(row=1, column=1, padx=10, pady=(10, 0), sticky="ew")

        interval_type_label = ctk.CTkLabel(edit_window, text="Interval Type:")
        interval_type_label.grid(row=2, column=0, padx=10, pady=(10, 0), sticky="w")

        interval_type_var = ctk.StringVar(value="Fixed")
        interval_type_option = ctk.CTkOptionMenu(
            edit_window,
            variable=interval_type_var,
            values=["Fixed", "Randomized"],
            command=lambda _: self.toggle_interval_type_edit(
                interval_type_var,
                fixed_interval_entry,
                min_interval_label,
                min_interval_entry,
                max_interval_label,
                max_interval_entry,
            ),
        )
        interval_type_option.grid(row=2, column=1, padx=10, pady=(10, 0), sticky="ew")

        fixed_interval_label = ctk.CTkLabel(edit_window, text="Interval (seconds):")
        fixed_interval_label.grid(row=3, column=0, padx=10, pady=(10, 0), sticky="w")

        fixed_interval_var = ctk.StringVar(
            value=str(action.interval) if action.interval else "1.0"
        )
        fixed_interval_entry = ctk.CTkEntry(
            edit_window, textvariable=fixed_interval_var
        )
        fixed_interval_entry.grid(row=3, column=1, padx=10, pady=(10, 0), sticky="ew")

        min_interval_label = ctk.CTkLabel(edit_window, text="Min Interval (seconds):")
        min_interval_label.grid(row=4, column=0, padx=10, pady=(10, 0), sticky="w")

        min_interval_var = ctk.StringVar(
            value=str(action.min_interval) if action.min_interval else "0.5"
        )
        min_interval_entry = ctk.CTkEntry(edit_window, textvariable=min_interval_var)
        min_interval_entry.grid(row=4, column=1, padx=10, pady=(10, 0), sticky="ew")

        max_interval_label = ctk.CTkLabel(edit_window, text="Max Interval (seconds):")
        max_interval_label.grid(row=5, column=0, padx=10, pady=(10, 0), sticky="w")

        max_interval_var = ctk.StringVar(
            value=str(action.max_interval) if action.max_interval else "1.5"
        )
        max_interval_entry = ctk.CTkEntry(edit_window, textvariable=max_interval_var)
        max_interval_entry.grid(row=5, column=1, padx=10, pady=(10, 0), sticky="ew")

        self.toggle_interval_type_edit(
            interval_type_var,
            fixed_interval_entry,
            min_interval_label,
            min_interval_entry,
            max_interval_label,
            max_interval_entry,
        )

        button_frame = ctk.CTkFrame(edit_window)
        button_frame.grid(row=6, column=0, columnspan=2, pady=20)

        save_button = ctk.CTkButton(
            button_frame,
            text="Save",
            command=lambda: self.on_save_edit_action(
                index,
                action_type_var,
                value_var,
                interval_type_var,
                fixed_interval_var,
                min_interval_var,
                max_interval_var,
                edit_window,
            ),
        )
        save_button.grid(row=0, column=0, padx=10)

        delete_button = ctk.CTkButton(
            button_frame,
            text="Delete Action",
            fg_color="red",
            command=lambda: self.on_delete_action(index, edit_window),
        )
        delete_button.grid(row=0, column=1, padx=10)

        cancel_button = ctk.CTkButton(
            button_frame, text="Cancel", command=edit_window.destroy
        )
        cancel_button.grid(row=0, column=2, padx=10)

        edit_window.grid_columnconfigure(1, weight=1)

    def on_mouse_button_selected_edit(
        self, button: str, value_var: ctk.StringVar, window: ctk.CTkToplevel
    ):
        value_var.set(button)
        window.destroy()

    def update_edit_value_widget(
        self,
        action_type_var: ctk.StringVar,
        value_var: ctk.StringVar,
        edit_window: ctk.CTkToplevel,
    ):
        if action_type_var.get() == "mouse_click":

            button_window = ctk.CTkToplevel(edit_window)
            button_window.title("Select Mouse Button")
            button_window.geometry("300x200")
            button_window.resizable(False, False)
            button_window.grab_set()

            label = ctk.CTkLabel(button_window, text="Select Mouse Button:")
            label.pack(pady=10)

            mouse_buttons = ["left", "right", "middle"]
            for btn in mouse_buttons:
                btn_widget = ctk.CTkButton(
                    button_window,
                    text=btn.capitalize(),
                    command=lambda b=btn: self.on_mouse_button_selected_edit(
                        b, value_var, button_window
                    ),
                )
                btn_widget.pack(pady=5)
        else:

            pass

    def toggle_randomization_edit(
        self,
        parent: ctk.CTkToplevel,
        randomize_var: ctk.BooleanVar,
        interval_label: ctk.CTkLabel,
        interval_entry: ctk.CTkEntry,
        min_interval_label: ctk.CTkLabel,
        min_interval_entry: ctk.CTkEntry,
        max_interval_label: ctk.CTkLabel,
        max_interval_entry: ctk.CTkEntry,
    ):
        if randomize_var.get():

            min_interval_label.grid()
            min_interval_entry.grid()
            max_interval_label.grid()
            max_interval_entry.grid()

            interval_label.grid_remove()
            interval_entry.grid_remove()
        else:

            min_interval_label.grid_remove()
            min_interval_entry.grid_remove()
            max_interval_label.grid_remove()
            max_interval_entry.grid_remove()

            interval_label.grid()
            interval_entry.grid()

    def get_children_with_label(
        self, parent: ctk.CTkToplevel, label_text: str
    ) -> List[ctk.CTkLabel]:
        """Helper method to retrieve widgets based on their label text within a parent widget."""
        return [
            child
            for child in parent.winfo_children()
            if isinstance(child, ctk.CTkLabel) and child.cget("text") == label_text
        ]

    def on_save_edit_action(
        self,
        index: int,
        action_type_var: ctk.StringVar,
        value_var: ctk.StringVar,
        interval_type_var: ctk.StringVar,
        fixed_interval_var: ctk.StringVar,
        min_interval_var: ctk.StringVar,
        max_interval_var: ctk.StringVar,
        edit_window: ctk.CTkToplevel,
    ):
        try:
            new_action_type = action_type_var.get()
            new_value = value_var.get()
            interval_type = interval_type_var.get()

            if interval_type == "Randomized":
                min_interval = float(min_interval_var.get())
                max_interval = float(max_interval_var.get())
                if min_interval <= 0 or max_interval <= 0:
                    raise ValueError("Intervals must be positive numbers.")
                if min_interval >= max_interval:
                    raise ValueError("Min Interval must be less than Max Interval.")
                self.action_sequence[index].interval = None
                self.action_sequence[index].min_interval = min_interval
                self.action_sequence[index].max_interval = max_interval
            else:
                interval = float(fixed_interval_var.get())
                if interval <= 0:
                    raise ValueError("Interval must be a positive number.")
                self.action_sequence[index].interval = interval
                self.action_sequence[index].min_interval = None
                self.action_sequence[index].max_interval = None

            self.action_sequence[index].action_type = new_action_type
            self.action_sequence[index].value = new_value

            self.update_action_sequence_display()
            edit_window.destroy()
            print(f"Action {index + 1} edited.")
        except ValueError as ve:
            self.show_message("Error", str(ve))
        except Exception as e:
            self.show_message("Error", f"Failed to edit action: {e}")

    def on_delete_action(self, index: int, edit_window: ctk.CTkToplevel):
        del self.action_sequence[index]
        self.update_action_sequence_display()
        edit_window.destroy()
        print(f"Action {index + 1} deleted.")

    def clear_log(self):
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", ctk.END)
        self.log_textbox.configure(state="disabled")
        print("Action log cleared.")

    def delete_action_sequence(self):
        self.action_sequence = []
        self.update_action_sequence_display()

    def show_message(self, title: str, message: str):
        message_dialog = CustomMessageBox(self.root, title=title, message=message)
        self.root.wait_window(message_dialog)
        print(f"{title}: {message}")

    def change_appearance_mode(self, mode: str):
        ctk.set_appearance_mode(mode)
        print(f"Appearance mode changed to: {mode}")

    def on_closing(self):
        self.stop_action()
        self.root.destroy()
        print("Application closed.")

    def start_recording(self):
        if self.recording:
            self.show_message("Info", "Recording is already in progress.")
            return

        self.recorded_events.clear()
        self.recording = True
        self.last_event_time = time.time()
        self.start_record_button.configure(state=ctk.DISABLED)
        self.stop_record_button.configure(state=ctk.NORMAL)
        self.keyboard_listener = keyboard.hook(self.on_keyboard_event)
        self.mouse_listener = mouse.hook(self.on_mouse_event)

        print("Recording started.")

    def stop_recording(self):
        if not self.recording:
            self.show_message("Info", "No recording is in progress.")
            return

        self.recording = False
        self.start_record_button.configure(state=ctk.NORMAL)
        self.stop_record_button.configure(state=ctk.DISABLED)

        if self.keyboard_listener:
            keyboard.unhook(self.keyboard_listener)
            self.keyboard_listener = None
        if self.mouse_listener:
            mouse.unhook(self.mouse_listener)
            self.mouse_listener = None

        self.show_message(
            "Recording Stopped", "Recording has stopped. Processing events."
        )
        print("Recording stopped.")

        self.process_recorded_events()

    def on_keyboard_event(self, event):
        if not self.recording:
            return
        try:
            if event.event_type == "down":
                key = event.name
                if key:
                    current_time = time.time()
                    if self.last_event_time is not None:
                        interval = current_time - self.last_event_time
                    else:
                        interval = 0
                    self.last_event_time = current_time
                    self.recorded_events.append(("key_press", key, interval))
                    print(
                        f"Recorded key press: {key} with interval {interval:.2f} seconds"
                    )
        except Exception as e:
            print(f"Error recording keyboard event: {e}")

    def on_mouse_event(self, event):
        if not self.recording:
            return
        try:
            if isinstance(event, mouse.ButtonEvent):
                if event.event_type == "down":
                    try:
                        x, y = pyautogui.position()
                    except Exception as pos_error:
                        print(f"Failed to get mouse position: {pos_error}")
                        return

                    self.root.update_idletasks()

                    excluded_widgets = [
                        self.stop_record_button,
                    ]

                    if any(
                        self.is_click_on_widget(widget, x, y)
                        for widget in excluded_widgets
                    ):
                        print("Excluded widget clicked. Ignoring this click.")
                        return

                    button = event.button
                    current_time = time.time()
                    if self.last_event_time is not None:
                        interval = current_time - self.last_event_time
                    else:
                        interval = 0
                    self.last_event_time = current_time
                    self.recorded_events.append(("mouse_click", button, interval))
                    print(
                        f"Recorded mouse click: {button} with interval {interval:.2f} seconds"
                    )
        except Exception as e:
            print(f"Error recording mouse event: {e}")

    def process_recorded_events(self):
        for event_type, value, interval in self.recorded_events:
            action = Action(
                action_type=event_type,
                value=value,
                interval=interval,
                target_window=self.target_window,
            )
            self.action_sequence.append(action)
        self.update_action_sequence_display()
        print("Recorded events processed and added to the action sequence.")

    def toggle_interval_type_edit(
        self,
        interval_type_var: ctk.StringVar,
        fixed_interval_entry: ctk.CTkEntry,
        min_interval_label: ctk.CTkLabel,
        min_interval_entry: ctk.CTkEntry,
        max_interval_label: ctk.CTkLabel,
        max_interval_entry: ctk.CTkEntry,
    ):
        if interval_type_var.get() == "Randomized":

            min_interval_label.grid()
            min_interval_entry.grid()
            max_interval_label.grid()
            max_interval_entry.grid()

            fixed_interval_entry.grid_remove()
        else:

            min_interval_label.grid_remove()
            min_interval_entry.grid_remove()
            max_interval_label.grid_remove()
            max_interval_entry.grid_remove()

            fixed_interval_entry.grid()

    def is_click_on_widget(
        self, widget: ctk.CTkButton, click_x: int, click_y: int
    ) -> bool:
        widget_x = widget.winfo_rootx()
        widget_y = widget.winfo_rooty()
        widget_width = widget.winfo_width()
        widget_height = widget.winfo_height()

        within_x = widget_x <= click_x <= widget_x + widget_width
        within_y = widget_y <= click_y <= widget_y + widget_height
        return within_x and within_y

    def refresh_window_list(self):
        if not self.use_target_window_var.get():

            return

        try:
            windows = [window for window in gw.getAllWindows() if window.title]
            if not windows:
                self.window_option_menu.configure(values=["No active windows found"])
                self.selected_window_var.set("No active windows found")
                print("No active windows found.")
                return

            window_list = [
                f"{idx + 1}. {window.title}" for idx, window in enumerate(windows)
            ]
            self.window_handles = [window._hWnd for window in windows]
            self.window_option_menu.configure(values=window_list)
            self.selected_window_var.set("Select a window")
            print("Window list refreshed.")
        except Exception as e:
            self.show_message("Error", f"Failed to refresh window list: {e}")
            print(f"Failed to refresh window list: {e}")

    def set_target_window(self):
        if not self.use_target_window_var.get():
            self.show_message("Info", "Targeted window actions are disabled.")
            print(
                "Attempted to set target window while targeted window actions are disabled."
            )
            return

        selected_option = self.selected_window_var.get()
        print(f"Selected option: {selected_option}")
        if selected_option in ["Select a window", "No active windows found"]:
            self.target_window = None
            self.show_message("Info", "Please select a valid window.")
            print("No valid window selected.")
            self.update_all_actions_target_window(None)
            return

        try:

            index = int(selected_option.split(".")[0]) - 1
            window_handle = self.window_handles[index]
            print(f"Selected window handle: {window_handle}")
            self.target_window = window_handle
            self.show_message(
                "Success", f"Target window set to: {selected_option.split('. ', 1)[1]}"
            )
            print(f"Target window set to: {selected_option.split('. ', 1)[1]}")
            self.update_all_actions_target_window(window_handle)
        except Exception as e:
            self.show_message("Error", f"Failed to set target window: {e}")
            print(f"Failed to set target window: {e}")
            self.target_window = None
            self.selected_window_var.set("Select a window")
            self.update_all_actions_target_window(None)

    def bring_window_to_foreground(self):
        try:
            hwnd = self.target_window
            if hwnd:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                print(f"Window {hwnd} brought to foreground.")
        except Exception as e:
            print(f"Failed to bring window to foreground: {e}")

    def toggle_target_window_options(self):
        if self.use_target_window_var.get():

            self.refresh_windows_button.configure(state=ctk.NORMAL)
            self.window_option_menu.configure(state="normal")
            self.refresh_window_list()
        else:

            self.refresh_windows_button.configure(state=ctk.DISABLED)
            self.window_option_menu.configure(state="disabled")
            self.selected_window_var.set("Select a window")
            self.target_window = None
            print("Targeted window actions disabled.")
            self.update_all_actions_target_window(None)

    def update_log(self, action: Action):
        if not isinstance(action, Action):
            print(f"Invalid action type: {type(action)}. Expected Action object.")
            return

        current_time = time.strftime("%H:%M:%S")
        if action.action_type == "key_press":
            log_message = f"[{current_time}] Pressed key: {action.value}"
        elif action.action_type == "mouse_click":
            log_message = f"[{current_time}] Clicked mouse button: {action.value}"
        else:
            log_message = f"[{current_time}] Unknown action: {action.action_type}"

        if action.interval is not None:
            interval = action.interval
            log_message += f" | Interval: {interval:.2f}s"
        elif action.min_interval is not None and action.max_interval is not None:
            interval = action.get_interval()
            log_message += f" | Interval: Random ({action.min_interval:.2f}s - {action.max_interval:.2f}s) (used {interval:.2f}s)"
        else:
            interval = 1.0

        self.log_textbox.configure(state="normal")
        self.log_textbox.insert(ctk.END, log_message + "\n")
        self.log_textbox.see(ctk.END)
        self.log_textbox.configure(state="disabled")
        print(log_message)

    def update_all_actions_target_window(self, new_target_window: Optional[int]):
        for action in self.action_sequence:
            action.target_window = new_target_window
        print(f"All actions updated to target_window: {new_target_window}")


if __name__ == "__main__":
    root = ctk.CTk()
    app = AutoAction(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()
