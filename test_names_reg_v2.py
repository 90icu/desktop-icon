import win32api
import win32con
import winreg

def get_monitor_name_from_key(key_path):
    try:
        # key_path is like \Registry\Machine\System\CurrentControlSet\Control\Class\...
        # Remove \Registry\Machine\ prefix
        prefix = "\\Registry\\Machine\\"
        if key_path.startswith(prefix):
            key_path = key_path[len(prefix):]
            
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
            try:
                # Try DriverDesc
                desc, _ = winreg.QueryValueEx(key, "DriverDesc")
                return desc
            except:
                pass
    except Exception as e:
        print(f"Reg error for {key_path}: {e}")
    return None

def test_monitor_names_reg_v2():
    print("Enumerating Monitors...")
    try:
        monitors = win32api.EnumDisplayMonitors()
        for i, (handle, _, rect) in enumerate(monitors):
            print(f"\nMonitor {i}:")
            info = win32api.GetMonitorInfo(handle)
            device_name = info['Device']
            
            try:
                monitor_dev = win32api.EnumDisplayDevices(device_name, 0)
                print(f"  DeviceKey: {monitor_dev.DeviceKey}")
                
                real_name = get_monitor_name_from_key(monitor_dev.DeviceKey)
                print(f"  Registry Name: {real_name}")
                
            except Exception as e:
                print(f"  Error: {e}")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    test_monitor_names_reg_v2()
