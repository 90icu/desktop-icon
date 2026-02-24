import win32api
import win32con

def test_monitor_names():
    print("Enumerating Monitors...")
    try:
        monitors = win32api.EnumDisplayMonitors()
        for i, (handle, _, rect) in enumerate(monitors):
            print(f"\nMonitor {i}: Handle {handle}")
            info = win32api.GetMonitorInfo(handle)
            device_name = info['Device'] # Adapter name, e.g. \\.\DISPLAY1
            is_primary = info['Flags'] & win32con.MONITORINFOF_PRIMARY
            print(f"  Device: {device_name}")
            print(f"  Primary: {bool(is_primary)}")
            
            # Try to get Monitor Name
            try:
                # EnumDisplayDevices(device_name, 0) gets the first monitor attached to this adapter output
                monitor_dev = win32api.EnumDisplayDevices(device_name, 0)
                print(f"  Monitor Name: {monitor_dev.DeviceString}")
                print(f"  Monitor ID: {monitor_dev.DeviceID}")
                print(f"  Monitor Key: {monitor_dev.DeviceKey}")
            except Exception as e:
                print(f"  Error getting monitor device: {e}")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    test_monitor_names()
