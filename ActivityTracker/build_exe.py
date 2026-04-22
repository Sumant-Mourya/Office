"""Build ActivityTracker.exe from main.py using PyInstaller.

Output:
    ActivityTracker.exe in the project root directory.
"""

import os
import subprocess
import sys


def main() -> int:
    root_dir = os.path.dirname(os.path.abspath(__file__))
    main_py = os.path.join(root_dir, "main.py")
    credentials = os.path.join(root_dir, "credentials.json")

    if not os.path.exists(main_py):
        print("main.py not found.")
        return 1

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name",
        "ActivityTracker",
        "--distpath",
        root_dir,
        "--workpath",
        os.path.join(root_dir, "build"),
        "--specpath",
        root_dir,
        "--collect-all",
        "nicegui",
        "--hidden-import",
        "pynput.keyboard._win32",
        "--hidden-import",
        "pynput.mouse._win32",
    ]

    if os.path.exists(credentials):
        cmd.extend(["--add-data", f"{credentials};."])

    cmd.append(main_py)

    print("Running:", " ".join(cmd))
    result = subprocess.run(cmd, cwd=root_dir)
    if result.returncode != 0:
        return result.returncode

    spec_file = os.path.join(root_dir, "ActivityTracker.spec")
    if os.path.exists(spec_file):
        try:
            os.remove(spec_file)
        except OSError:
            pass

    print("Build complete: ActivityTracker.exe")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
