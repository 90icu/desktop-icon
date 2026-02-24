import ctypes
import struct
import win32gui
import win32con
import win32api
import win32process
import winreg
import json
import os
import time
import logging

# Setup logging - DISABLED FILE LOGGING as requested
# logging.basicConfig(
#     filename='desktop_manager.log', 
#     level=logging.DEBUG, 
#     format='%(asctime)s - %(levelname)s - %(message)s',
#     encoding='utf-8'
# )
# Use a custom logger that stores last error for UI retrieval
class MemoryLogger:
    def __init__(self):
        self.last_error = None
        
    def debug(self, msg): pass
    def info(self, msg): pass
    def warning(self, msg): pass
    
    def error(self, msg):
        self.last_error = msg
        print(f"ERROR: {msg}") # Still print to console just in case
        
    def critical(self, msg):
        self.last_error = msg
        print(f"CRITICAL: {msg}")

logging = MemoryLogger()

# Constants
LVM_FIRST = 0x1000
LVM_GETITEMCOUNT = LVM_FIRST + 4
LVM_GETITEMTEXTW = LVM_FIRST + 115
LVM_GETITEMPOSITION = LVM_FIRST + 16
LVM_SETITEMPOSITION32 = LVM_FIRST + 49
LVM_GETITEMSPACING = LVM_FIRST + 53
LVS_AUTOARRANGE = 0x0100
GWL_STYLE = -16

PROCESS_ALL_ACCESS = 0x1F0FFF
MEM_COMMIT = 0x1000
MEM_RESERVE = 0x2000
MEM_RELEASE = 0x8000
PAGE_READWRITE = 0x04

class DesktopManager:
    def __init__(self):
        self.hwnd = self._get_desktop_listview()
        if not self.hwnd:
            logging.error("Could not find Desktop ListView handle")
            raise Exception("Could not find Desktop ListView handle")
        
        pid = win32process.GetWindowThreadProcessId(self.hwnd)[1]
        self.process = ctypes.windll.kernel32.OpenProcess(PROCESS_ALL_ACCESS, False, pid)
        if not self.process:
            logging.error("Could not open Desktop process")
            raise Exception("Could not open Desktop process")
        
        # Allocate shared buffer for move operations (8 bytes for x,y)
        self.move_buffer = ctypes.windll.kernel32.VirtualAllocEx(self.process, 0, 8, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE)
        if not self.move_buffer:
            error = ctypes.GetLastError()
            logging.error(f"Failed to allocate move_buffer in __init__. Error: {error}")
            # Try to continue? If this fails, move_icon will fail anyway.
        else:
            logging.info(f"Allocated move_buffer at {hex(self.move_buffer)}")

        logging.info(f"Initialized DesktopManager. HWND: {self.hwnd}, PID: {pid}")

    def _get_desktop_listview(self):
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

    def _read_memory(self, address, size):
        buffer = ctypes.create_string_buffer(size)
        bytes_read = ctypes.c_size_t()
        ctypes.windll.kernel32.ReadProcessMemory(self.process, ctypes.c_void_p(address), buffer, size, ctypes.byref(bytes_read))
        return buffer.raw

    def _write_memory(self, address, data):
        bytes_written = ctypes.c_size_t()
        ctypes.windll.kernel32.WriteProcessMemory(self.process, ctypes.c_void_p(address), data, len(data), ctypes.byref(bytes_written))

    def get_icon_spacing(self):
        spacing = win32gui.SendMessage(self.hwnd, LVM_GETITEMSPACING, 0, 0)
        width = spacing & 0xFFFF
        height = (spacing >> 16) & 0xFFFF
        logging.info(f"Icon Spacing: {width}x{height}")
        return width, height

    def get_monitors(self):
        monitors = []
        try:
            for i, handle in enumerate(win32api.EnumDisplayMonitors()):
                h_monitor, h_dc, (left, top, right, bottom) = handle
                info = win32api.GetMonitorInfo(h_monitor)
                monitors.append({
                    "index": i,
                    "handle": int(h_monitor),
                    "rect": info['Monitor'],
                    "work": info['Work'],
                    "device": info['Device'],
                    "is_primary": (info['Flags'] & win32con.MONITORINFOF_PRIMARY) != 0
                })
        except Exception as e:
            logging.error(f"EnumDisplayMonitors failed: {e}")
        return monitors

    def move_icon(self, index, x, y):
        if not self.move_buffer:
            logging.error("move_buffer is not allocated")
            return False

        try:
            data = struct.pack('ii', x, y)
            self._write_memory(self.move_buffer, data)
            win32gui.SendMessage(self.hwnd, LVM_SETITEMPOSITION32, index, self.move_buffer)
            return True
        except Exception as e:
            logging.error(f"Failed to move icon {index}: {e}")
            return False

    def get_icons(self):
        count = win32gui.SendMessage(self.hwnd, LVM_GETITEMCOUNT, 0, 0)
        logging.info(f"Found {count} icons on desktop.")
        
        lvitem_size = 128 
        text_buffer_size = 1024 
        
        remote_mem = ctypes.windll.kernel32.VirtualAllocEx(self.process, 0, lvitem_size + text_buffer_size, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE)
        remote_point = ctypes.windll.kernel32.VirtualAllocEx(self.process, 0, 8, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE)
        
        if not remote_mem or not remote_point:
            error_code = ctypes.GetLastError()
            logging.error(f"Failed to allocate memory in desktop process. Error: {error_code}")
            return [], (100, 100)

        spacing_x, spacing_y = self.get_icon_spacing()
        monitors = self.get_monitors()
        monitor_handle_map = {int(m['handle']): m for m in monitors}
        
        icons = []

        try:
            for i in range(count):
                # 1. Get Text
                lvitem_buffer = ctypes.create_string_buffer(lvitem_size)
                struct.pack_into("I", lvitem_buffer, 0, 0x0001) # mask = LVIF_TEXT
                struct.pack_into("i", lvitem_buffer, 4, i)      # iItem
                struct.pack_into("i", lvitem_buffer, 8, 0)      # iSubItem
                text_ptr_addr = remote_mem + lvitem_size
                struct.pack_into("Q", lvitem_buffer, 24, text_ptr_addr) 
                struct.pack_into("i", lvitem_buffer, 32, 512)   # cchTextMax

                self._write_memory(remote_mem, lvitem_buffer.raw)
                win32gui.SendMessage(self.hwnd, LVM_GETITEMTEXTW, i, remote_mem)
                
                text_raw = self._read_memory(text_ptr_addr, 1024)
                name = text_raw.decode('utf-16').split('\x00')[0]

                # 2. Get Position
                win32gui.SendMessage(self.hwnd, LVM_GETITEMPOSITION, i, remote_point)
                point_raw = self._read_memory(remote_point, 8)
                phys_x, phys_y = struct.unpack('ii', point_raw)

                # 3. Find Monitor
                try:
                    h_mon = win32api.MonitorFromPoint((phys_x, phys_y), win32con.MONITOR_DEFAULTTONEAREST)
                    h_mon_int = int(h_mon)
                    
                    if h_mon_int in monitor_handle_map:
                        m = monitor_handle_map[h_mon_int]
                        monitor_idx = m['index']
                        monitor_left = m['rect'][0]
                        monitor_top = m['rect'][1]
                    else:
                        monitor_idx = 0 
                        monitor_left = monitors[0]['rect'][0]
                        monitor_top = monitors[0]['rect'][1]
                            
                except Exception as e:
                    logging.warning(f"MonitorFromPoint failed for icon {name}: {e}")
                    monitor_idx = 0
                    monitor_left = 0
                    monitor_top = 0
                
                # Calculate Relative Position
                rel_x = phys_x - monitor_left
                rel_y = phys_y - monitor_top
                
                col = round(rel_x / spacing_x)
                row = round(rel_y / spacing_y)

                icons.append({
                    "name": name,
                    "x": phys_x,
                    "y": phys_y,
                    "monitor": monitor_idx,
                    "row": row,
                    "col": col
                })
        except Exception as e:
            logging.error(f"Error during get_icons loop: {e}")
        finally:
            if remote_mem:
                ctypes.windll.kernel32.VirtualFreeEx(self.process, remote_mem, 0, MEM_RELEASE)
            if remote_point:
                ctypes.windll.kernel32.VirtualFreeEx(self.process, remote_point, 0, MEM_RELEASE)
            # DO NOT CLOSE HANDLE HERE. It belongs to the class instance.


        return icons, (spacing_x, spacing_y)

    def is_point_on_screen(self, x, y, monitors):
        for m in monitors:
            r = m['rect']
            if r[0] <= x < r[2] and r[1] <= y < r[3]:
                return True
        return False

    def restore_icons(self, saved_icons, saved_monitors=None, progress_callback=None):
        logging.info("Starting restore_icons...")
        
        # 0. Check Auto Arrange
        style = win32gui.GetWindowLong(self.hwnd, GWL_STYLE)
        if style & LVS_AUTOARRANGE:
            logging.warning("Auto Arrange detected! Disabling...")
            win32gui.SetWindowLong(self.hwnd, GWL_STYLE, style & ~LVS_AUTOARRANGE)
            # time.sleep(0.5)

        # 1. Get CURRENT icons
        current_icons, current_spacing = self.get_icons()
        if not current_icons:
            logging.error("No icons found on desktop! Aborting.")
            return 0
            
        current_map = {icon['name']: i for i, icon in enumerate(current_icons)}
        
        # 2. Match
        saved_names = set(icon['name'] for icon in saved_icons)
        current_names = set(current_map.keys())
        common_names = saved_names.intersection(current_names)
        
        logging.info(f"Matching Icons: {len(common_names)}")
        
        if len(common_names) == 0:
            logging.critical("NO MATCHING ICONS FOUND! Aborting restore.")
            return 0

        # 3. Setup Monitors
        monitors = self.get_monitors()
        current_spacing_x, current_spacing_y = current_spacing
        current_device_map = {m['device']: m for m in monitors}
        current_primary = next((m for m in monitors if m['is_primary']), monitors[0])
        
        monitor_mapping = {}
        if saved_monitors:
            for sm in saved_monitors:
                s_idx = sm['index']
                s_device = sm.get('device')
                target = current_primary
                
                if s_device and s_device in current_device_map:
                    target = current_device_map[s_device]
                elif sm.get('is_primary'):
                    target = current_primary
                elif s_idx < len(monitors):
                    target = monitors[s_idx]
                
                monitor_mapping[s_idx] = target
        
        # 4. Evacuate
        logging.info("Evacuating icons...")
        total_items = win32gui.SendMessage(self.hwnd, LVM_GETITEMCOUNT, 0, 0)
        # for i in range(total_items):
        #     self.move_icon(i, 20000, 0)
        
        # time.sleep(1.0)
        
        # 5. Restore Loop
        restored_count = 0
        
        # Sort saved_icons by Monitor -> Row -> Col to ensure consistent restore order (Top-Left to Bottom-Right)
        # If 'row'/'col' are missing or wrong, we fallback to x, y for sorting
        try:
            saved_icons.sort(key=lambda x: (x.get('monitor', 0), x.get('y', 0), x.get('x', 0)))
        except:
            logging.warning("Failed to sort icons, using default order")
            
        for saved in saved_icons:
            name = saved['name']
            if name not in current_map:
                continue
                
            idx = current_map[name]
            
            target_x = 0
            target_y = 0
            
            # --- Coordinate Logic ---
            if saved_monitors is None:
                # Legacy Mode / Old JSON
                if self.is_point_on_screen(saved['x'], saved['y'], monitors):
                    target_x = saved['x']
                    target_y = saved['y']
                    logging.info(f"Legacy {name}: Using absolute ({target_x}, {target_y})")
                else:
                    # Off-screen or weird. Map to Primary.
                    # Recalculate safe col/row
                    col = saved.get('col', 0)
                    row = saved.get('row', 0)
                    
                    # Heuristic: If col is huge (>30) and we are mapping to single monitor,
                    # it probably means it was on a second monitor.
                    # We should shift it.
                    max_cols = (current_primary['rect'][2] - current_primary['rect'][0]) // current_spacing_x
                    if max_cols < 1: max_cols = 1
                    
                    if col >= max_cols:
                        col = col % max_cols # Wrap around
                        
                    target_x = current_primary['rect'][0] + int(col * current_spacing_x)
                    target_y = current_primary['rect'][1] + int(row * current_spacing_y)
                    logging.info(f"Legacy {name}: Remapped to Primary ({target_x}, {target_y})")
            else:
                # New Mode
                saved_mon_idx = saved.get('monitor', 0)
                target_monitor = monitor_mapping.get(saved_mon_idx, current_primary)
                
                col = saved['col']
                row = saved['row']
                
                target_x = target_monitor['rect'][0] + int(col * current_spacing_x)
                target_y = target_monitor['rect'][1] + int(row * current_spacing_y)
                
                # Bounds check
                if not self.is_point_on_screen(target_x, target_y, monitors):
                     logging.warning(f"{name} target ({target_x}, {target_y}) is off-screen. Forcing to Primary.")
                     target_x = current_primary['rect'][0] + int(col * current_spacing_x)
                     target_y = current_primary['rect'][1] + int(row * current_spacing_y)

            # Move
            if self.move_icon(idx, target_x, target_y):
                restored_count += 1
            else:
                logging.error(f"Failed to move {name}")
                
            # Slow down
            # time.sleep(0.1)
            
            if progress_callback:
                try:
                    progress_callback(restored_count, len(saved_icons))
                except:
                    pass
            
        # 6. Refresh
        win32gui.InvalidateRect(self.hwnd, None, True)
        win32gui.UpdateWindow(self.hwnd)
        
        return restored_count

    def close(self):
        if self.move_buffer:
             ctypes.windll.kernel32.VirtualFreeEx(self.process, self.move_buffer, 0, MEM_RELEASE)
             self.move_buffer = None

        if self.process:
            ctypes.windll.kernel32.CloseHandle(self.process)
            self.process = None

def get_current_layout_data():
    """获取当前桌面布局数据，不保存到文件"""
    dm = DesktopManager()
    try:
        icons, spacing = dm.get_icons()
        # Use rich monitor info
        monitors = get_monitors_info()
        icons.sort(key=lambda x: (x['monitor'], x['row'], x['col']))
        
        return {
            "version": "3.6",
            "timestamp": time.time(),
            "monitors": monitors,
            "spacing": spacing,
            "icons": icons
        }
    finally:
        dm.close()

def restore_from_data(data, progress_callback=None):
    """从数据对象恢复布局"""
    dm = DesktopManager()
    try:
        saved_monitors = data.get('monitors', None)
        return dm.restore_icons(data['icons'], saved_monitors, progress_callback)
    finally:
        dm.close()

def get_monitor_registry_name(device_key):
    """
    Try to get the friendly name of the monitor from the registry.
    device_key looks like: \Registry\Machine\System\CurrentControlSet\Control\Class\{...}\0001
    """
    try:
        prefix = "\\Registry\\Machine\\"
        if device_key.startswith(prefix):
            key_path = device_key[len(prefix):]
        else:
            key_path = device_key
            
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
            try:
                # Try DriverDesc first
                desc, _ = winreg.QueryValueEx(key, "DriverDesc")
                return desc
            except:
                pass
    except:
        pass
    return None

def get_monitors_info():
    """
    Get information about all connected monitors.
    Returns a list of dicts:
    [{
        "index": int,
        "device": str,
        "rect": (left, top, right, bottom),
        "resolution": (width, height),
        "position": (x, y),
        "refresh_rate": int,
        "is_primary": bool,
        "name": str,
        "device_id": str
    }, ...]
    """
    monitors = []
    try:
        # Get all monitor handles
        monitor_handles = win32api.EnumDisplayMonitors()
        
        for i, (handle, _, rect) in enumerate(monitor_handles):
            monitor_info = win32api.GetMonitorInfo(handle)
            device_name = monitor_info['Device']
            is_primary = bool(monitor_info['Flags'] & win32con.MONITORINFOF_PRIMARY)
            
            # Get display settings for refresh rate
            try:
                settings = win32api.EnumDisplaySettings(device_name, win32con.ENUM_CURRENT_SETTINGS)
                refresh_rate = settings.DisplayFrequency
            except Exception as e:
                logging.error(f"Error getting settings for {device_name}: {e}")
                refresh_rate = 60 # Default fallback
            
            # Get Monitor Name and ID
            name = "Unknown Monitor"
            device_id = ""
            try:
                # EnumDisplayDevices(device_name, 0) gets the first monitor attached to this adapter output
                monitor_dev = win32api.EnumDisplayDevices(device_name, 0)
                device_id = monitor_dev.DeviceID
                
                # Try registry name first
                reg_name = get_monitor_registry_name(monitor_dev.DeviceKey)
                if reg_name:
                    name = reg_name
                else:
                    name = monitor_dev.DeviceString
                    
                # If name is Generic PnP, try to append manufacturer ID from DeviceID
                # DeviceID example: MONITOR\BOE0B7D\{UUID}\0001
                if "Generic" in name and device_id.startswith("MONITOR\\"):
                    try:
                        parts = device_id.split("\\")
                        if len(parts) > 1:
                            mfg_id = parts[1]
                            name += f" ({mfg_id})"
                    except:
                        pass
                        
            except Exception as e:
                logging.error(f"Error getting monitor device info: {e}")
                
            monitors.append({
                "index": i,
                "device": device_name,
                "rect": rect, # (left, top, right, bottom)
                "work_area": monitor_info['Work'],
                "resolution": (rect[2] - rect[0], rect[3] - rect[1]),
                "position": (rect[0], rect[1]),
                "refresh_rate": refresh_rate,
                "is_primary": is_primary,
                "name": name,
                "device_id": device_id
            })
            
    except Exception as e:
        logging.error(f"Error enumerating monitors: {e}")
        
    return monitors

def save_layout(filename="desktop_layout.json"):
    data = get_current_layout_data()
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    return len(data['icons'])

def restore_layout(filename="desktop_layout.json", progress_callback=None):
    if not os.path.exists(filename):
        return 0
        
    with open(filename, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    return restore_from_data(data, progress_callback)
