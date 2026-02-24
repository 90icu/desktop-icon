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

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


class _Logger:
    """最小日志器，仅保留最后一条错误供 UI 展示。"""
    def __init__(self):
        self.last_error = None
    def debug(self, msg): pass
    def info(self, msg): pass
    def warning(self, msg): pass
    def error(self, msg):
        self.last_error = msg
    def critical(self, msg):
        self.last_error = msg


logging = _Logger()

LVM_FIRST            = 0x1000
LVM_GETITEMCOUNT     = LVM_FIRST + 4
LVM_GETITEMTEXTW     = LVM_FIRST + 115
LVM_GETITEMPOSITION  = LVM_FIRST + 16
LVM_SETITEMPOSITION32 = LVM_FIRST + 49
LVM_GETITEMSPACING   = LVM_FIRST + 53
LVS_AUTOARRANGE      = 0x0100
GWL_STYLE            = -16

PROCESS_ALL_ACCESS = 0x1F0FFF
MEM_COMMIT    = 0x1000
MEM_RESERVE   = 0x2000
MEM_RELEASE   = 0x8000
PAGE_READWRITE = 0x04


class DesktopManager:
    def __init__(self):
        self.hwnd = self._get_desktop_listview()
        if not self.hwnd:
            raise Exception("Could not find Desktop ListView handle")

        pid = win32process.GetWindowThreadProcessId(self.hwnd)[1]
        self.process = ctypes.windll.kernel32.OpenProcess(PROCESS_ALL_ACCESS, False, pid)
        if not self.process:
            raise Exception("Could not open Desktop process")

        self.move_buffer = ctypes.windll.kernel32.VirtualAllocEx(
            self.process, 0, 8, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE)
        if not self.move_buffer:
            logging.error(f"Failed to allocate move_buffer. Error: {ctypes.GetLastError()}")

    def _get_desktop_listview(self):
        hwnd_progman = win32gui.FindWindow("Progman", "Program Manager")
        hwnd_shell = win32gui.FindWindowEx(hwnd_progman, 0, "SHELLDLL_DefView", None)
        hwnd_listview = win32gui.FindWindowEx(hwnd_shell, 0, "SysListView32", None)

        if not hwnd_listview:
            def _enum(hwnd, results):
                if win32gui.GetClassName(hwnd) == "WorkerW":
                    shell = win32gui.FindWindowEx(hwnd, 0, "SHELLDLL_DefView", None)
                    if shell:
                        lv = win32gui.FindWindowEx(shell, 0, "SysListView32", None)
                        if lv:
                            results.append(lv)
                return True
            results = []
            win32gui.EnumWindows(_enum, results)
            if results:
                hwnd_listview = results[0]

        return hwnd_listview

    def _read_memory(self, address, size):
        buffer = ctypes.create_string_buffer(size)
        bytes_read = ctypes.c_size_t()
        ctypes.windll.kernel32.ReadProcessMemory(
            self.process, ctypes.c_void_p(address), buffer, size, ctypes.byref(bytes_read))
        return buffer.raw

    def _write_memory(self, address, data):
        bytes_written = ctypes.c_size_t()
        ctypes.windll.kernel32.WriteProcessMemory(
            self.process, ctypes.c_void_p(address), data, len(data), ctypes.byref(bytes_written))

    def get_icon_spacing(self):
        spacing = win32gui.SendMessage(self.hwnd, LVM_GETITEMSPACING, 0, 0)
        return spacing & 0xFFFF, (spacing >> 16) & 0xFFFF

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
            self._write_memory(self.move_buffer, struct.pack('ii', x, y))
            win32gui.SendMessage(self.hwnd, LVM_SETITEMPOSITION32, index, self.move_buffer)
            return True
        except Exception as e:
            logging.error(f"Failed to move icon {index}: {e}")
            return False

    def get_icons(self):
        count = win32gui.SendMessage(self.hwnd, LVM_GETITEMCOUNT, 0, 0)

        lvitem_size = 128
        text_buffer_size = 1024

        remote_mem = ctypes.windll.kernel32.VirtualAllocEx(
            self.process, 0, lvitem_size + text_buffer_size, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE)
        remote_point = ctypes.windll.kernel32.VirtualAllocEx(
            self.process, 0, 8, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE)

        if not remote_mem or not remote_point:
            logging.error(f"Failed to allocate memory. Error: {ctypes.GetLastError()}")
            return [], (100, 100)

        spacing_x, spacing_y = self.get_icon_spacing()
        monitors = self.get_monitors()
        monitor_handle_map = {int(m['handle']): m for m in monitors}

        # LVM_GETITEMPOSITION 返回虚拟桌面坐标（原点=虚拟桌面左上角）
        # GetMonitorInfo 返回屏幕坐标（主显示器左上角为原点）
        # 转换：screen = lvm + virtual_offset
        virtual_left = min(m['rect'][0] for m in monitors) if monitors else 0
        virtual_top  = min(m['rect'][1] for m in monitors) if monitors else 0

        icons = []
        try:
            for i in range(count):
                lvitem_buffer = ctypes.create_string_buffer(lvitem_size)
                struct.pack_into("I", lvitem_buffer, 0, 0x0001)
                struct.pack_into("i", lvitem_buffer, 4, i)
                struct.pack_into("i", lvitem_buffer, 8, 0)
                text_ptr_addr = remote_mem + lvitem_size
                struct.pack_into("Q", lvitem_buffer, 24, text_ptr_addr)
                struct.pack_into("i", lvitem_buffer, 32, 512)

                self._write_memory(remote_mem, lvitem_buffer.raw)
                win32gui.SendMessage(self.hwnd, LVM_GETITEMTEXTW, i, remote_mem)

                text_raw = self._read_memory(text_ptr_addr, 1024)
                name = text_raw.decode('utf-16').split('\x00')[0]

                win32gui.SendMessage(self.hwnd, LVM_GETITEMPOSITION, i, remote_point)
                lvm_x, lvm_y = struct.unpack('ii', self._read_memory(remote_point, 8))
                screen_x = lvm_x + virtual_left
                screen_y = lvm_y + virtual_top

                monitor_device = ""
                try:
                    h_mon = win32api.MonitorFromPoint(
                        (screen_x, screen_y), win32con.MONITOR_DEFAULTTONEAREST)
                    h_mon_int = int(h_mon)
                    if h_mon_int in monitor_handle_map:
                        m = monitor_handle_map[h_mon_int]
                        monitor_idx  = m['index']
                        monitor_left = m['rect'][0]
                        monitor_top  = m['rect'][1]
                        monitor_device = m.get('device', '')
                    else:
                        monitor_idx  = 0
                        monitor_left = monitors[0]['rect'][0]
                        monitor_top  = monitors[0]['rect'][1]
                        monitor_device = monitors[0].get('device', '')
                except Exception as e:
                    logging.warning(f"MonitorFromPoint failed for {name}: {e}")
                    monitor_idx  = 0
                    monitor_left = virtual_left
                    monitor_top  = virtual_top

                icons.append({
                    "name": name,
                    "x": screen_x,
                    "y": screen_y,
                    "monitor": monitor_idx,
                    "monitor_device": monitor_device,
                    "col": round((screen_x - monitor_left) / spacing_x) if spacing_x else 0,
                    "row": round((screen_y - monitor_top)  / spacing_y) if spacing_y else 0,
                })
        except Exception as e:
            logging.error(f"Error in get_icons loop: {e}")
        finally:
            if remote_mem:
                ctypes.windll.kernel32.VirtualFreeEx(self.process, remote_mem, 0, MEM_RELEASE)
            if remote_point:
                ctypes.windll.kernel32.VirtualFreeEx(self.process, remote_point, 0, MEM_RELEASE)

        return icons, (spacing_x, spacing_y)

    def is_point_on_screen(self, x, y, monitors):
        return any(m['rect'][0] <= x < m['rect'][2] and m['rect'][1] <= y < m['rect'][3]
                   for m in monitors)

    def restore_icons(self, saved_icons, saved_monitors=None, progress_callback=None):
        style = win32gui.GetWindowLong(self.hwnd, GWL_STYLE)
        if style & LVS_AUTOARRANGE:
            win32gui.SetWindowLong(self.hwnd, GWL_STYLE, style & ~LVS_AUTOARRANGE)

        current_icons, current_spacing = self.get_icons()
        if not current_icons:
            logging.error("No icons found on desktop! Aborting.")
            return 0

        current_map = {icon['name']: i for i, icon in enumerate(current_icons)}
        common_names = set(ic['name'] for ic in saved_icons) & set(current_map.keys())
        if not common_names:
            logging.critical("NO MATCHING ICONS FOUND! Aborting restore.")
            return 0

        monitors = self.get_monitors()
        current_spacing_x, current_spacing_y = current_spacing
        current_device_map = {m['device']: m for m in monitors}
        current_primary = next((m for m in monitors if m['is_primary']), monitors[0])

        # LVM 坐标转换偏移
        virtual_left = min(m['rect'][0] for m in monitors) if monitors else 0
        virtual_top  = min(m['rect'][1] for m in monitors) if monitors else 0

        monitor_mapping = {}
        if saved_monitors:
            for sm in saved_monitors:
                s_device = sm.get('device')
                if s_device and s_device in current_device_map:
                    target = current_device_map[s_device]
                elif sm.get('is_primary'):
                    target = current_primary
                elif sm['index'] < len(monitors):
                    target = monitors[sm['index']]
                else:
                    target = current_primary
                monitor_mapping[sm['index']] = target

        try:
            saved_icons.sort(key=lambda x: (x.get('monitor', 0), x.get('row', 0), x.get('col', 0)))
        except Exception:
            pass

        restored_count = 0
        for saved in saved_icons:
            name = saved['name']
            if name not in current_map:
                continue

            idx = current_map[name]

            if saved_monitors is None:
                target_x = saved['x']
                target_y = saved['y']
            else:
                target_monitor = monitor_mapping.get(saved.get('monitor', 0), current_primary)
                col = saved['col']
                row = saved['row']

                target_x = target_monitor['rect'][0] + int(col * current_spacing_x)
                target_y = target_monitor['rect'][1] + int(row * current_spacing_y)

                if not self.is_point_on_screen(target_x, target_y, monitors):
                    logging.warning(f"{name} off-screen, forcing to primary.")
                    target_x = current_primary['rect'][0] + int(col * current_spacing_x)
                    target_y = current_primary['rect'][1] + int(row * current_spacing_y)

                # 屏幕坐标 → LVM 虚拟桌面坐标
                target_x -= virtual_left
                target_y -= virtual_top

            if self.move_icon(idx, target_x, target_y):
                restored_count += 1
            else:
                logging.error(f"Failed to move {name}")

            if progress_callback:
                try:
                    progress_callback(restored_count, len(saved_icons))
                except Exception:
                    pass

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



def _get_monitor_registry_name(device_key):
    try:
        prefix = "\\Registry\\Machine\\"
        key_path = device_key[len(prefix):] if device_key.startswith(prefix) else device_key
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
            try:
                desc, _ = winreg.QueryValueEx(key, "DriverDesc")
                return desc
            except Exception:
                pass
    except Exception:
        pass
    return None


def get_monitors_info():
    """返回所有显示器的详细信息列表。"""
    monitors = []
    try:
        for i, (handle, _, rect) in enumerate(win32api.EnumDisplayMonitors()):
            info = win32api.GetMonitorInfo(handle)
            device_name = info['Device']
            is_primary = bool(info['Flags'] & win32con.MONITORINFOF_PRIMARY)

            try:
                settings = win32api.EnumDisplaySettings(device_name, win32con.ENUM_CURRENT_SETTINGS)
                refresh_rate = settings.DisplayFrequency
            except Exception:
                refresh_rate = 60

            name = "Unknown Monitor"
            device_id = ""
            try:
                monitor_dev = win32api.EnumDisplayDevices(device_name, 0)
                device_id = monitor_dev.DeviceID
                reg_name = _get_monitor_registry_name(monitor_dev.DeviceKey)
                name = reg_name if reg_name else monitor_dev.DeviceString
                if "Generic" in name and device_id.startswith("MONITOR\\"):
                    parts = device_id.split("\\")
                    if len(parts) > 1:
                        name += f" ({parts[1]})"
            except Exception:
                pass

            monitors.append({
                "index": i,
                "device": device_name,
                "rect": rect,
                "work_area": info['Work'],
                "resolution": (rect[2] - rect[0], rect[3] - rect[1]),
                "position": (rect[0], rect[1]),
                "refresh_rate": refresh_rate,
                "is_primary": is_primary,
                "name": name,
                "device_id": device_id,
            })
    except Exception as e:
        logging.error(f"Error enumerating monitors: {e}")
    return monitors


def get_current_layout_data():
    """获取当前桌面布局数据。"""
    dm = DesktopManager()
    try:
        icons, spacing = dm.get_icons()
        monitors = get_monitors_info()

        device_to_vis_idx = {m['device']: m['index'] for m in monitors}
        for icon in icons:
            device = icon.get('monitor_device', '')
            if device and device in device_to_vis_idx:
                icon['monitor'] = device_to_vis_idx[device]

        icons.sort(key=lambda x: (x['monitor'], x['row'], x['col']))

        return {
            "version": "3.6",
            "timestamp": time.time(),
            "monitors": monitors,
            "spacing": spacing,
            "icons": icons,
        }
    finally:
        dm.close()


def restore_from_data(data, progress_callback=None):
    """从数据对象恢复布局。"""
    dm = DesktopManager()
    try:
        return dm.restore_icons(data['icons'], data.get('monitors'), progress_callback)
    finally:
        dm.close()


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
