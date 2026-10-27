import os
import subprocess
import sys


def main():
    project_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(project_dir)

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name",
        "HotelExpenseTracker",
        "--collect-all",
        "streamlit",
        "--copy-metadata",
        "streamlit",
        "--collect-all",
        "pandas",
        "--collect-all",
        "numpy",
        "--collect-all",
        "plotly",
        "--collect-all",
        "openpyxl",
        "--collect-all",
        "pymupdf",
    ]

    for source in ("app.py", "parser.py", "matcher.py", "exporter.py", "sr_parser.py", "ui_components.py"):
        cmd.extend(["--add-data", f"{source};."])

    for directory in ("data", "ui", "views"):
        cmd.extend(["--add-data", f"{directory};{directory}"])

    cmd.extend([
        "app_launcher.py",
    ])

    print("Building the portable Windows executable...")
    result = subprocess.run(cmd)
    if result.returncode == 0:
        print("\nBuild success.")
        print("Output folder: dist/")
        print("Executable: dist/HotelExpenseTracker.exe")
    else:
        print("\nBuild failed.")
        print("Run this script on Windows after installing the project and PyInstaller.")
        sys.exit(result.returncode)


if __name__ == "__main__":
    main()
