import os
import subprocess
import sys


def main():
    project_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(project_dir)

    print("Building Windows EXE for the Hotel Expense Tracker...")
    print("This packages the Streamlit app into a single executable.")

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--onefile",
        "--noconsole",
        "app_launcher.py",
    ]

    result = subprocess.run(cmd)
    if result.returncode == 0:
        print("\nBuild success.")
        print("Output folder: dist/")
        print("Executable: dist/app_launcher.exe")
    else:
        print("\nBuild failed.")
        print("Install PyInstaller first: pip install pyinstaller")


if __name__ == "__main__":
    main()
