import tkinter as tk
from tkinter import ttk
import win32gui
import win32con

# Store all windows in a list
window_list = []

def enum_windows():
    window_list.clear()
    def callback(hwnd, extra):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
            window_list.append((win32gui.GetWindowText(hwnd), hwnd))
    win32gui.EnumWindows(callback, None)

def refresh_window_list():
    enum_windows()
    window_dropdown['values'] = [title for title, hwnd in window_list]
    if window_list:
        window_dropdown.current(0)

def set_always_on_top():
    selected_title = selected_window.get()
    for title, hwnd in window_list:
        if title == selected_title:
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST,
                                  0, 0, 0, 0,
                                  win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            status_label.config(text=f"'{title}' pinned on top!")
            return

def remove_always_on_top():
    selected_title = selected_window.get()
    for title, hwnd in window_list:
        if title == selected_title:
            win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST,
                                  0, 0, 0, 0,
                                  win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            status_label.config(text=f"'{title}' unpinned.")
            return

# Create GUI
root = tk.Tk()
root.title("Always On Top Tool")
root.geometry("400x200")

selected_window = tk.StringVar()

ttk.Label(root, text="Select a window:").pack(pady=5)

window_dropdown = ttk.Combobox(root, textvariable=selected_window, width=50)
window_dropdown.pack()

ttk.Button(root, text="Refresh List", command=refresh_window_list).pack(pady=5)

ttk.Button(root, text="Pin (Always on Top)", command=set_always_on_top).pack(pady=5)
ttk.Button(root, text="Unpin", command=remove_always_on_top).pack(pady=5)

status_label = ttk.Label(root, text="")
status_label.pack(pady=10)

refresh_window_list()  # Load initial list
root.mainloop()
