# Windows packaging notes for this app

This project is a Streamlit app. The easiest way to make a Windows `.exe` is to package a small launcher that starts Streamlit.

## 1) Install builder

```bash
pip install pyinstaller
```

## 2) Build the EXE

From the project folder:

```bash
pyinstaller --onefile --noconsole app_launcher.py
```

This will generate a file in the `dist` folder, such as:

```text
dist/app_launcher.exe
```

## 3) Runtime notes

The app expects a local data directory, usually with month folders such as:

```text
Business Review/
  8.AGUSTUS/
  7.JULY/
```

You can set a custom folder before launch:

```powershell
set HOTEL_DATA_DIR=C:\Business Review
app_launcher.exe
```

or choose the folder manually inside the app if you later add a folder picker.

## 4) Why this approach

- Fastest way to package the app without rewriting the UI
- Keeps your current Streamlit dashboard intact
- Suitable for internal business use

## 5) For a more native Windows app later

If you want a fully native desktop feel later, the next step is to rebuild the interface using Electron or PyQt, but that would be a larger migration.
