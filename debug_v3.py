import ctypes
import struct
import win32gui
import win32con
import win32process
import win32api
import json
import os
import time

# Constants
LVM_FIRST = 0x1000
LVM_GETITEMCOUNT = LVM_FIRST + 4
LVM_GETITEMTEXTW = LVM_FIRST + 115
LVM_GETITEMPOSITION = LVM_FIRST + 16
LVM_SETITEMPOSITION32 = LVM_FIRST + 49
LVS_AUTOARRANGE = 0x0100
GWL_STYLE = -16

PROCESS_ALL_ACCESS = 0x1F0FFF
MEM_COMMIT = 0x1000
MEM_RESERVE = 0x2000
MEM_RELEASE = 0x8000
PAGE_READWRITE = 0x04

def get_desktop_listview():
    hwnd_progman = win32gui.FindWindow("Progman", "Program Manager")
    hwnd_shell = win32gui.FindWindowEx(hwnd_progman, 0, "SHELLDLL_DefView", None)
    hwnd_listview = win32gui.FindWindowEx(hwnd_shell, 0, "SysListView32", None)

    if not hwnd_listview:
        def enum_windows_callback(hwnd, result_list):
            if win32gui.GetClassName(hwnd) == "WorkerW":
                hwnd_shell = win32gui.FindWindowEx(hwnd, 0, "SHELLDLL_DefView", None)
                if hwnd_shell:
                    hwnd_lv = win32gui.FindWindowEx(hwnd_shell, 0, "SysListView32", None)
                    if hwnd_lv:
                        result_list.append(hwnd_lv)
            return True

        results = []
        win32gui.EnumWindows(enum_windows_callback, results)
        if results:
            hwnd_listview = results[0]
            
    return hwnd_listview

def analyze():
    print("--- Desktop Analysis ---")
    hwnd = get_desktop_listview()
    if not hwnd:
        print("ERROR: Could not find Desktop ListView")
        return

    print(f"ListView HWND: {hwnd}")
    
    # Check Styles
    style = win32gui.GetWindowLong(hwnd, GWL_STYLE)
    print(f"Window Style: {hex(style)}")
    if style & LVS_AUTOARRANGE:
        print("WARNING: Auto Arrange is ON! (Icons will snap back)")
    else:
        print("Auto Arrange is OFF.")

    # Get Count
    count = win32gui.SendMessage(hwnd, LVM_GETITEMCOUNT, 0, 0)
    print(f"Icon Count: {count}")

    if count == 0:
        return

    # Process Memory
    pid = win32process.GetWindowThreadProcessId(hwnd)[1]
    process = ctypes.windll.kernel32.OpenProcess(PROCESS_ALL_ACCESS, False, pid)
    
    # Read first 5 icons
    lvitem_size = 128 
    text_buffer_size = 512
    remote_mem = ctypes.windll.kernel32.VirtualAllocEx(process, 0, lvitem_size + text_buffer_size, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE)
    remote_point = ctypes.windll.kernel32.VirtualAllocEx(process, 0, 8, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE)
    
    print("\n--- First 5 Icons ---")
    try:
        for i in range(min(5, count)):
            # Name
            lvitem_buffer = ctypes.create_string_buffer(lvitem_size)
            struct.pack_into("I", lvitem_buffer, 0, 0x0001) # LVIF_TEXT
            struct.pack_into("i", lvitem_buffer, 4, i)
            struct.pack_into("i", lvitem_buffer, 8, 0)
            text_ptr_addr = remote_mem + lvitem_size
            struct.pack_into("Q", lvitem_buffer, 24, text_ptr_addr)
            struct.pack_into("i", lvitem_buffer, 32, 256)

            bytes_written = ctypes.c_size_t()
            ctypes.windll.kernel32.WriteProcessMemory(process, ctypes.c_void_p(remote_mem), lvitem_buffer.raw, len(lvitem_buffer.raw), ctypes.byref(bytes_written))
            win32gui.SendMessage(hwnd, LVM_GETITEMTEXTW, i, remote_mem)
            
            buffer = ctypes.create_string_buffer(512)
            bytes_read = ctypes.c_size_t()
            ctypes.windll.kernel32.ReadProcessMemory(process, ctypes.c_void_p(text_ptr_addr), buffer, 512, ctypes.byref(bytes_read))
            name = buffer.raw.decode('utf-16').split('\x00')[0]
            
            # Pos
            win32gui.SendMessage(hwnd, LVM_GETITEMPOSITION, i, remote_point)
            point_raw = ctypes.create_string_buffer(8)
            ctypes.windll.kernel32.ReadProcessMemory(process, ctypes.c_void_p(remote_point), point_raw, 8, ctypes.byref(bytes_read))
            x, y = struct.unpack('ii', point_raw.raw)
            
            print(f"Icon {i}: '{name}' at ({x}, {y})")
            
    finally:
        ctypes.windll.kernel32.VirtualFreeEx(process, remote_mem, 0, MEM_RELEASE)
        ctypes.windll.kernel32.VirtualFreeEx(process, remote_point, 0, MEM_RELEASE)
        ctypes.windll.kernel32.CloseHandle(process)

    # Check JSON
    print("\n--- JSON File Check ---")
    if os.path.exists("desktop_layout.json"):
        try:
            with open("desktop_layout.json", 'r', encoding='utf-8') as f:
                data = json.load(f)
                print(f"JSON Version: {data.get('version')}")
                print(f"Saved Icons: {len(data.get('icons', []))}")
                if data.get('icons'):
                    print(f"First Saved Icon: {data['icons'][0]['name']} at ({data['icons'][0]['x']}, {data['icons'][0]['y']})")
        except Exception as e:
            print(f"Error reading JSON: {e}")
    else:
        print("desktop_layout.json not found.")

if __name__ == "__main__":
    analyze()
