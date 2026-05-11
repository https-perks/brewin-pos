import os
import sys
import time
import shutil
from pathlib import Path

def main():
    if len(sys.argv) != 4:
        # Arguments are: updater.exe, current_exe_path, new_exe_temp_path, version
        sys.exit(1)

    current_exe = Path(sys.argv[1])
    new_exe = Path(sys.argv[2])
    version = sys.argv[3]

    # Wait for the main app to exit completely
    time.sleep(1.0)

    for _ in range(50):   # up to 5 seconds
        try:
            os.remove(current_exe)
            break
        except PermissionError:
            time.sleep(0.1)

    # Replace the old EXE with the new version
    try:
        shutil.move(str(new_exe), str(current_exe))
    except Exception as e:
        # Write emergency log
        logpath = Path(os.environ.get("LOCALAPPDATA", "")) / "BrewInsPOS" / "update_error.log"
        with open(logpath, "a", encoding="utf-8") as f:
            f.write(f"Update move failed: {e}\n")
        sys.exit(1)

    # Relaunch the updated application
    try:
        os.startfile(str(current_exe))
    except Exception:
        pass

if __name__ == "__main__":
    main()
