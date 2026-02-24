
import ctypes
import win32gui
import win32api
import win32con
import struct
import win32process

def get_dpi_for_monitor(h_monitor):
    try:
        # MDT_EFFECTIVE_DPI = 0
        dpi_x = ctypes.c_uint()
        dpi_y = ctypes.c_uint()
        ctypes.windll.shcore.GetDpiForMonitor(
            int(h_monitor), 
            0, 
            ctypes.byref(dpi_x), 
            ctypes.byref(dpi_y)
        )
        return dpi_x.value
    except Exception as e:
        print(f"Failed to get DPI: {e}")
        return 96

def main():
    print("=== DPI DIAGNOSIS ===")
    
    # 1. Set DPI Aware
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2) # Per Monitor
        print("DPI Awareness: Per Monitor (2)")
    except:
        print("Failed to set DPI Awareness")

    # 2. Virtual Screen (Physical)
    vx = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
    vy = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
    print(f"Virtual Screen (Physical?): Left={vx}, Top={vy}")

    # 3. Monitors
    print("\n--- Monitors ---")
    monitors = []
    for i, handle in enumerate(win32api.EnumDisplayMonitors()):
        h_monitor, h_dc, (left, top, right, bottom) = handle
        dpi = get_dpi_for_monitor(h_monitor)
        scale = dpi / 96.0
        
        info = win32api.GetMonitorInfo(h_monitor)
        rect = info['Monitor']
        
        print(f"Monitor {i}: Handle={h_monitor}")
        print(f"  Physical Rect: {rect}")
        print(f"  DPI: {dpi} (Scale: {scale:.2f})")
        
        # Calculate Logical Rect (approximate)
        # Note: This is tricky because coordinates are continuous
        # We need to map physical coordinates to logical space
        
        monitors.append({
            "handle": h_monitor,
            "rect": rect,
            "scale": scale
        })

    # 4. Desktop Icons (ListView is Logical Coordinates)
    print("\n--- Desktop Icons (First 5) ---")
    progman = win32gui.FindWindow("Progman", None)
    hwnd_shell = win32gui.FindWindowEx(progman, 0, "SHELLDLL_DefView", None)
    hwnd_lv = win32gui.FindWindowEx(hwnd_shell, 0, "SysListView32", None)
    
    if not hwnd_lv:
        # Try WorkerW
        workerw = win32gui.FindWindowEx(0, 0, "WorkerW", None)
        while workerw:
            hwnd_shell = win32gui.FindWindowEx(workerw, 0, "SHELLDLL_DefView", None)
            if hwnd_shell:
                hwnd_lv = win32gui.FindWindowEx(hwnd_shell, 0, "SysListView32", None)
                break
            workerw = win32gui.FindWindowEx(0, workerw, "WorkerW", None)

    if not hwnd_lv:
        print("Could not find ListView")
        return

    print(f"ListView HWND: {hwnd_lv}")
    
    # Read first 5 icons
    count = win32gui.SendMessage(hwnd_lv, 0x1000 + 4, 0, 0) # LVM_GETITEMCOUNT
    pid = win32process.GetWindowThreadProcessId(hwnd_lv)[1]
    process = ctypes.windll.kernel32.OpenProcess(0x1F0FFF, False, pid)
    
    pt_buf = ctypes.windll.kernel32.VirtualAllocEx(process, 0, 8, 0x1000, 0x04)
    
    for i in range(min(count, 5)):
        win32gui.SendMessage(hwnd_lv, 0x1000 + 16, i, pt_buf) # LVM_GETITEMPOSITION
        
        data = ctypes.create_string_buffer(8)
        bytes_read = ctypes.c_size_t()
        ctypes.windll.kernel32.ReadProcessMemory(process, pt_buf, data, 8, ctypes.byref(bytes_read))
        x, y = struct.unpack('ii', data)
        
        print(f"Icon {i}: LV_Pos=({x}, {y})")
        
        # Analyze which monitor it falls into
        # We need to convert LV_Pos (Logical) to Physical to match Monitor Rects?
        # OR convert Monitor Rects to Logical?
        
        # Hypothesis: LV_Pos is relative to Virtual Screen Top-Left in LOGICAL pixels.
        # But Virtual Screen Top-Left in LOGICAL pixels depends on the DPI of the primary monitor?
        # Let's try to map it.
        
    ctypes.windll.kernel32.VirtualFreeEx(process, pt_buf, 0, 0x8000)
    ctypes.windll.kernel32.CloseHandle(process)

if __name__ == "__main__":
    main()
