
import win32process
import win32gui
import win32api
import win32con
import ctypes
import struct
import time

def move_test():
    print("=== MOVE TEST ===")
    
    # 1. Setup
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except:
        pass
    
    progman = win32gui.FindWindow("Progman", None)
    hwnd_shell = win32gui.FindWindowEx(progman, 0, "SHELLDLL_DefView", None)
    hwnd_lv = win32gui.FindWindowEx(hwnd_shell, 0, "SysListView32", None)
    
    if not hwnd_lv:
        workerw = win32gui.FindWindowEx(0, 0, "WorkerW", None)
        while workerw:
            hwnd_shell = win32gui.FindWindowEx(workerw, 0, "SHELLDLL_DefView", None)
            if hwnd_shell:
                hwnd_lv = win32gui.FindWindowEx(hwnd_shell, 0, "SysListView32", None)
                break
            workerw = win32gui.FindWindowEx(0, workerw, "WorkerW", None)
            
    print(f"ListView HWND: {hwnd_lv}")
    
    # 2. Get Virtual Screen Info
    v_left = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
    print(f"Virtual Screen Left: {v_left}")
    
    # 3. Pick the first icon
    pid = win32process.GetWindowThreadProcessId(hwnd_lv)[1]
    process = ctypes.windll.kernel32.OpenProcess(0x1F0FFF, False, pid)
    
    pt_buf = ctypes.windll.kernel32.VirtualAllocEx(process, 0, 8, 0x1000, 0x04)
    
    # Read current pos
    win32gui.SendMessage(hwnd_lv, 0x1000 + 16, 0, pt_buf) # GETPOSITION
    data = ctypes.create_string_buffer(8)
    ctypes.windll.kernel32.ReadProcessMemory(process, pt_buf, data, 8, 0)
    orig_x, orig_y = struct.unpack('ii', data)
    print(f"Icon 0 Original Pos: ({orig_x}, {orig_y})")
    
    # 4. Test Move to 0
    print("Moving Icon 0 to (0, 0)...")
    packed = struct.pack('ii', 0, 0)
    ctypes.windll.kernel32.WriteProcessMemory(process, pt_buf, packed, 8, 0)
    win32gui.SendMessage(hwnd_lv, 0x1000 + 49, 0, pt_buf) # SETPOSITION32
    
    # Wait user to see
    time.sleep(2)
    
    # 5. Test Move to -2560 (if v_left is -2560, this should be same as 0?)
    # Wait, if Origin is V_Left, then 0 is V_Left.
    # If Origin is Primary, then -2560 is V_Left.
    
    print(f"Moving Icon 0 to ({abs(v_left)}, 0) -> Should be Primary Left if Origin=V_Left")
    packed = struct.pack('ii', abs(v_left), 0)
    ctypes.windll.kernel32.WriteProcessMemory(process, pt_buf, packed, 8, 0)
    win32gui.SendMessage(hwnd_lv, 0x1000 + 49, 0, pt_buf)
    
    time.sleep(2)
    
    # 6. Restore
    print(f"Restoring to ({orig_x}, {orig_y})")
    packed = struct.pack('ii', orig_x, orig_y)
    ctypes.windll.kernel32.WriteProcessMemory(process, pt_buf, packed, 8, 0)
    win32gui.SendMessage(hwnd_lv, 0x1000 + 49, 0, pt_buf)
    
    # Refresh
    win32gui.InvalidateRect(hwnd_lv, None, True)
    
    ctypes.windll.kernel32.VirtualFreeEx(process, pt_buf, 0, 0x8000)
    ctypes.windll.kernel32.CloseHandle(process)

if __name__ == "__main__":
    move_test()
