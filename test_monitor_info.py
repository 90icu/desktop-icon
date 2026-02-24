import win32api
import win32con

def get_monitors_info():
    monitors = []
    try:
        # Get all monitor handles
        monitor_handles = win32api.EnumDisplayMonitors()
        
        for handle, _, rect in monitor_handles:
            monitor_info = win32api.GetMonitorInfo(handle)
            device_name = monitor_info['Device']
            
            # Get display settings for refresh rate
            try:
                settings = win32api.EnumDisplaySettings(device_name, win32con.ENUM_CURRENT_SETTINGS)
                refresh_rate = settings.DisplayFrequency
            except Exception as e:
                print(f"Error getting settings for {device_name}: {e}")
                refresh_rate = 60 # Default fallback
                
            monitors.append({
                "device": device_name,
                "rect": rect, # (left, top, right, bottom)
                "work_area": monitor_info['Work'],
                "resolution": (rect[2] - rect[0], rect[3] - rect[1]),
                "position": (rect[0], rect[1]),
                "refresh_rate": refresh_rate
            })
            
    except Exception as e:
        print(f"Error enumerating monitors: {e}")
        
    return monitors

if __name__ == "__main__":
    infos = get_monitors_info()
    for m in infos:
        print(m)
