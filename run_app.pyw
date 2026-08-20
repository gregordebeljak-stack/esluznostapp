import sys
import os
import webbrowser
import time
import docx_engine  # PyInstaller will see this and bundle it
from streamlit.web import cli as stcli


def resolve_path(path):
    # PyInstaller creates a temp folder and stores path in _MEIPASS
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, path)
    return os.path.join(os.path.abspath(os.path.dirname(__file__)), path)


if __name__ == '__main__':
    # Point this to your actual Streamlit application script
    script_path = resolve_path('app.py')

    # Keep Streamlit server startup deterministic and user-friendly for packaged windows launchers.
    sys.argv = ['streamlit', 'run', script_path, '--global.developmentMode=false']

    # Give the server a brief moment to initialize before opening the default address.
    # This keeps the launcher visible in the background without showing a console window.
    timer = 2
    try:
        if hasattr(sys, 'frozen'):
            timer = 4
    except Exception:
        pass

    # Start Streamlit in a GUI-friendly way. pythonw.exe has no console attached.
    sys.exit(stcli.main())
