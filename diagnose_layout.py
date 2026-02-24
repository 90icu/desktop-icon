
import json
import win32api
import win32con
import win32gui
import ctypes
import os

def get_monitors():
    monitors = []
    for i, handle in enumerate(win32api.EnumDisplayMonitors()):
        h_monitor, h_dc, (left, top, right, bottom) = handle
        info = win32api.GetMonitorInfo(h_monitor)
        monitors.append({
            "index": i,
            "handle": int(h_monitor),
            "rect": (left, top, right, bottom),
            "device": info['Device'],
            "is_primary": (info['Flags'] & win32con.MONITORINFOF_PRIMARY) != 0
        })
    return monitors

def diagnose():
    print("=== DIAGNOSTIC REPORT ===")
    
    # 1. System Metrics
    v_left = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
    v_top = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
    v_width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
    v_height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
    print(f"Virtual Screen: Left={v_left}, Top={v_top}, W={v_width}, H={v_height}")
    
    # 2. Monitors
    print("\n--- Current Monitors ---")
    monitors = get_monitors()
    for m in monitors:
        print(f"Monitor {m['index']}: {m['device']} | Rect={m['rect']} | Primary={m['is_primary']}")
        
    # 3. Load JSON
    json_path = "desktop_layout.json"
    if not os.path.exists(json_path):
        print("\n[ERROR] desktop_layout.json not found!")
        return
        
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        print(f"\n[ERROR] Failed to load JSON: {e}")
        return

    print(f"\n--- Saved Layout (Timestamp: {data.get('timestamp')}) ---")
    saved_monitors = data.get('monitors', [])
    print(f"Saved Monitors: {len(saved_monitors)}")
    for sm in saved_monitors:
        print(f"  Saved Mon {sm.get('index')}: {sm.get('device')} | Rect={sm.get('rect')} | Primary={sm.get('is_primary')}")
        
    icons = data.get('icons', [])
    print(f"\nSaved Icons: {len(icons)}")
    if icons:
        print("First 5 icons:")
        for icon in icons[:5]:
            print(f"  {icon['name']}: Mon={icon.get('monitor')} | Row={icon.get('row')} | Col={icon.get('col')}")
            
    # 4. Simulation
    print("\n--- Restore Simulation ---")
    # Simulate Mapping
    current_device_map = {m['device']: m for m in monitors}
    current_primary = next((m for m in monitors if m['is_primary']), monitors[0])
    
    monitor_mapping = {}
    for sm in saved_monitors:
        s_idx = sm['index']
        s_device = sm.get('device')
        
        target = None
        if s_device and s_device in current_device_map:
            target = current_device_map[s_device]
            match_type = "Device Name"
        elif sm.get('is_primary'):
            target = current_primary
            match_type = "Primary Flag"
        elif s_idx < len(monitors):
            target = monitors[s_idx]
            match_type = "Index"
        else:
            target = current_primary
            match_type = "Fallback"
            
        monitor_mapping[s_idx] = target
        print(f"Mapping Saved Mon {s_idx} -> Current {target['index']} ({target['device']}) via {match_type}")
        
    # Simulate Coordinate Calculation for first few icons
    current_spacing = win32gui.SendMessage(win32gui.FindWindow("Progman", "Program Manager"), 0x1000 + 53, 0, 0)
    # Note: This might fail if we don't get the right window, but let's try
    if current_spacing == 0:
         # Try grabbing spacing from data for simulation if we can't get real one easily without full class
         sx, sy = data.get('spacing', (100, 100))
    else:
         sx = current_spacing & 0xFFFF
         sy = (current_spacing >> 16) & 0xFFFF
         
    print(f"Using Spacing: {sx} x {sy}")
    
    for icon in icons[:5]:
        s_mon = icon.get('monitor', 0)
        target_mon = monitor_mapping.get(s_mon, current_primary)
        
        base_x = target_mon['rect'][0]
        base_y = target_mon['rect'][1]
        
        col = icon['col']
        row = icon['row']
        
        phys_x = base_x + int(col * sx)
        phys_y = base_y + int(row * sy)
        
        lv_x = phys_x - v_left
        lv_y = phys_y - v_top
        
        print(f"Icon '{icon['name']}': Saved(M{s_mon}, R{row}, C{col}) -> Target(M{target_mon['index']}) -> Phys({phys_x},{phys_y}) -> LV({lv_x},{lv_y})")

if __name__ == "__main__":
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except:
        pass
    diagnose()
