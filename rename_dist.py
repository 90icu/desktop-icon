import os
import glob

dist_dir = "dist"
target_name = "桌面图标管理.exe"
target_path = os.path.join(dist_dir, target_name)

# Find any exe in dist that is NOT the target name
files = glob.glob(os.path.join(dist_dir, "*.exe"))

for f in files:
    if os.path.basename(f) != target_name:
        try:
            # If target exists, remove it first
            if os.path.exists(target_path):
                os.remove(target_path)
            
            os.rename(f, target_path)
            print(f"Renamed {f} to {target_path}")
        except Exception as e:
            print(f"Error renaming {f}: {e}")
