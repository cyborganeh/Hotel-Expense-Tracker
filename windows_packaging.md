# Portable Windows executable

The app can be built as a single-file Windows executable that runs without an installer or administrator access. The executable contains Python, Streamlit, and the app code; your business data stays in its existing folder.

## Build on GitHub Actions

The repository workflow builds on a Windows runner, since PyInstaller cannot create Windows executables on Linux. Push the changes to GitHub, then open the repository's **Actions** tab and select **Build portable Windows EXE**. When the run finishes, download the `HotelExpenseTracker-windows-x64` artifact. GitHub provides it as a ZIP containing `HotelExpenseTracker.exe`.

The workflow also runs automatically on pushes to `main` that change app or packaging files.

## Build directly on Windows

With Python 3.11 installed, run these commands from the project folder:

```bash
python -m pip install . pyinstaller
python build_windows_exe.py
```

The executable is written to `dist/HotelExpenseTracker.exe`.

## Run

Copy the executable to a folder you can access and launch it. It opens the app in your browser at `http://localhost:8501`. No installation is required. Some managed workplaces may still block unsigned executables through application-control policy; an EXE cannot bypass that restriction.

The app expects your local data directory, usually with month folders such as:

```text
Business Review/
  8.AGUSTUS/
  7.JULY/
```

Set a custom folder before launching if needed:

```powershell
set HOTEL_DATA_DIR=C:\Business Review
HotelExpenseTracker.exe
```
