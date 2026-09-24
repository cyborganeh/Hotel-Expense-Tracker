import os
import subprocess
import sys

APP_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(APP_DIR)

cmd = [
    sys.executable,
    "-m",
    "streamlit",
    "run",
    "app.py",
    "--server.headless",
    "true",
    "--server.port",
    "8501",
]

print("Starting Hotel Expense Tracker...")
print("Open: http://localhost:8501")
subprocess.run(cmd, check=False)
