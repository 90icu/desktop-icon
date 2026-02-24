
import win32api
import win32con
import win32gui
import win32process
import ctypes
from ctypes import wintypes
import struct
import platform

def set_dpi_awareness():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # PROCESS_SYSTEM_DPI_AWARE
        print("DPI Awareness set to PROCESS_SYSTEM_DPI_AWARE")
    except Exception as e:
        print(f"Failed to set DPI awareness: {e}")

def get_virtual_screen_metrics():
    x = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
    y = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
    w = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
    h = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
    return x, y, w, h

def get_monitors():
    monitors = []
    for i, handle in enumerate(win32api.EnumDisplayMonitors()):
        h_monitor, h_dc, (left, top, right, bottom) = handle
        info = win32api.GetMonitorInfo(h_monitor)
        monitors.append({
            "handle": h_monitor,
            "rect": (left, top, right, bottom),
            "work": info['Work'],
            "is_primary": (info['Flags'] & win32con.MONITORINFOF_PRIMARY) != 0,
            "device": info['Device']
        })
    return monitors

def get_desktop_listview():
    progman = win32gui.FindWindow("Progman", None)
    if not progman:
        return None
    
    def enum_child_callback(hwnd, result_list):
        cls_name = win32gui.GetClassName(hwnd)
        if cls_name == "SysListView32":
            # Check if parent is SHELLDLL_DefView
            parent = win32gui.GetParent(hwnd)
            parent_cls = win32gui.GetClassName(parent)
            if parent_cls == "SHELLDLL_DefView":
                result_list.append(hwnd)
        return True

    # Try standard Progman path
    results = []
    win32gui.EnumChildWindows(progman, enum_child_callback, results)
    
    # If not found (common in Win10/11 due to WorkerW), try finding WorkerW
    if not results:
        workerw = win32gui.FindWindowEx(0, 0, "WorkerW", None)
        while workerw:
            win32gui.EnumChildWindows(workerw, enum_child_callback, results)
            if results: break
            workerw = win32gui.FindWindowEx(0, workerw, "WorkerW", None)
            
    return results[0] if results else None

def get_icon_info(hwnd):
    PROCESS_ALL_ACCESS = 0x1F0FFF
    pid = win32process.GetWindowThreadProcessId(hwnd)[1]
    process = win32api.OpenProcess(PROCESS_ALL_ACCESS, False, pid)
    
    LVM_FIRST = 0x1000
    LVM_GETITEMCOUNT = LVM_FIRST + 4
    LVM_GETITEMPOSITION = LVM_FIRST + 16
    
    count = win32gui.SendMessage(hwnd, LVM_GETITEMCOUNT, 0, 0)
    print(f"Found {count} icons.")
    
    # Allocate memory using ctypes
    MEM_COMMIT = 0x1000
    PAGE_READWRITE = 0x04
    MEM_RELEASE = 0x8000
    
    kernel32 = ctypes.windll.kernel32
    
    pt_buf = kernel32.VirtualAllocEx(int(process), 0, 8, MEM_COMMIT, PAGE_READWRITE)
    if not pt_buf:
        print("VirtualAllocEx failed")
        return []
        
    icons = []
    
    for i in range(min(count, 10)): # Check first 10 icons
        # Get Position
        win32gui.SendMessage(hwnd, LVM_GETITEMPOSITION, i, pt_buf)
        
        data = ctypes.create_string_buffer(8)
        bytes_read = ctypes.c_size_t(0)
        kernel32.ReadProcessMemory(int(process), pt_buf, data, 8, ctypes.byref(bytes_read))
        
        x, y = struct.unpack('ii', data)
        
        print(f"Icon {i}: Pos=({x}, {y})")
        icons.append((x,y))
        
    kernel32.VirtualFreeEx(int(process), pt_buf, 0, MEM_RELEASE)
    win32api.CloseHandle(process)
    return icons

def main():
    print("=== DEBUG COORDINATES SYSTEM ===")
    print(f"OS: {platform.system()} {platform.release()}")
    set_dpi_awareness()
    
    vx, vy, vw, vh = get_virtual_screen_metrics()
    print(f"Virtual Screen: Left={vx}, Top={vy}, Width={vw}, Height={vh}")
    
    monitors = get_monitors()
    print(f"Monitors Found: {len(monitors)}")
    for i, m in enumerate(monitors):
        print(f"  Monitor {i}: Rect={m['rect']}, Primary={m['is_primary']}, Device={m['device']}")
        
    hwnd = get_desktop_listview()
    if not hwnd:
        print("ERROR: Could not find Desktop SysListView32")
        return
        
    print(f"Desktop ListView HWND: {hwnd}")
    
    icons = get_icon_info(hwnd)
    
    print("\n--- ANALYSIS ---")
    for i, (x, y) in enumerate(icons):
        matched = False
        for m in monitors:
            ml, mt, mr, mb = m['rect']
            if ml <= x < mr and mt <= y < mb:
                print(f"Icon {i} ({x},{y}) is inside Monitor {m['device']} Rect ({ml},{mt},{mr},{mb})")
                matched = True
                break
        if not matched:
            print(f"Icon {i} ({x},{y}) is OUTSIDE all monitor rects!")

if __name__ == "__main__":
    main()
