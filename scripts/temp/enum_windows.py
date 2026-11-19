import ctypes
import ctypes.wintypes as wt

user32 = ctypes.windll.user32
pid_target = 13052

windows = []


def enum_proc(hwnd, _):
    pid = wt.DWORD()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    if pid.value == pid_target:
        cls = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, cls, 256)
        n = user32.GetWindowTextLengthW(hwnd)
        buf = ctypes.create_unicode_buffer(n + 1)
        user32.GetWindowTextW(hwnd, buf, n + 1)
        windows.append((hwnd, cls.value, buf.value, bool(user32.IsWindowVisible(hwnd))))
    return True


WNDENUMPROC = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)
user32.EnumWindows(WNDENUMPROC(enum_proc), 0)

for hwnd, cls, text, vis in windows:
    if text.strip() or "EXCEL" in cls.upper() or "VB" in cls.upper():
        print(f"hwnd={hwnd} visible={vis} class={cls!r} text={text!r}")
