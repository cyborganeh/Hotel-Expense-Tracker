import os
import sys

def main():
    app_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    app_path = os.path.join(app_dir, "app.py")
    os.chdir(app_dir)
    sys.path.insert(0, app_dir)

    from streamlit.web import cli as stcli

    sys.argv = ["streamlit", "run", app_path, "--server.port=8501"]
    stcli.main()


if __name__ == "__main__":
    main()
