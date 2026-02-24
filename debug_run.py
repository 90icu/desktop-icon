import traceback
import sys

print("--- DIAGNOSTIC RUN ---")
try:
    import desktop_manager
    print("Module imported.")
    print("Reading current desktop state...")
    desktop_manager.save_layout()
    print("\n--- DIAGNOSIS COMPLETE ---")
except Exception:
    traceback.print_exc()
