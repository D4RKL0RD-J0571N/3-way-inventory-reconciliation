"""
Standalone launcher for ReconcileAI.
Launches the Streamlit application directly from Python without relying on the global 'streamlit' CLI command.
Ideal for PyInstaller packaging and portable Python environments.
"""
import sys
import os
from pathlib import Path
import streamlit.web.cli as stcli

def main():
    base_dir = Path(__file__).resolve().parent
    os.chdir(base_dir)

    app_path = base_dir / "app.py"

    sys.argv = [
        "streamlit",
        "run",
        str(app_path),
        "--global.developmentMode=false",
        "--server.headless=false",
        "--browser.gatherUsageStats=false"
    ]
    sys.exit(stcli.main())

if __name__ == "__main__":
    main()
