import sys
import os
import docx_engine  # PyInstaller will see this and bundle it
from streamlit.web import cli as stcli

def resolve_path(path):
    # PyInstaller creates a temp folder and stores path in _MEIPASS
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, path)
    return os.path.join(os.path.abspath(os.path.dirname(__file__)), path)

if __name__ == '__main__':
    # Point this to your actual Streamlit application script
    script_path = resolve_path("app.py") 
    
    sys.argv = ["streamlit", "run", script_path, "--global.developmentMode=false"]
    sys.exit(stcli.main())