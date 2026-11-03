"""
Automated PyInstaller build script for ReconcileAI.
Packages the entire application into a portable, standalone executable directory (dist/ReconcileAI).
No Python installation required on the destination machine.
"""
import sys
import os
import shutil
from pathlib import Path
import PyInstaller.__main__
from PyInstaller.utils.hooks import copy_metadata, collect_data_files

BASE_DIR = Path(__file__).resolve().parent.parent

def build():
    print("=" * 70)
    print("  CREANDO EJECUTABLE PORTABLE DE RECONCILEAI (PyInstaller)")
    print("=" * 70)

    # Clean previous build artifacts
    build_dir = BASE_DIR / "build"
    dist_dir = BASE_DIR / "dist"
    if build_dir.exists():
        shutil.rmtree(build_dir, ignore_errors=True)

    # Collect necessary Streamlit, Pandas and Openpyxl data files & metadata
    datas = []
    datas += copy_metadata("streamlit")
    datas += collect_data_files("streamlit")
    datas += copy_metadata("openpyxl")
    datas += copy_metadata("pandas")

    # Add project resources
    datas.append((str(BASE_DIR / "app.py"), "."))
    datas.append((str(BASE_DIR / "src"), "src"))
    datas.append((str(BASE_DIR / ".streamlit"), ".streamlit"))
    datas.append((str(BASE_DIR / "docs"), "docs"))

    # Convert datas to PyInstaller CLI arguments
    data_args = []
    for src, dst in datas:
        data_args.extend(["--add-data", f"{src}{os.pathsep}{dst}"])

    pyinstaller_args = [
        str(BASE_DIR / "run_standalone.py"),
        "--name=ReconcileAI",
        "--onedir",
        "--noconfirm",
        "--clean",
        "--hidden-import=streamlit",
        "--hidden-import=openpyxl",
        "--hidden-import=pandas",
        "--hidden-import=altair",
        *data_args
    ]

    print("Ejecutando PyInstaller...")
    PyInstaller.__main__.run(pyinstaller_args)

    print("\n" + "=" * 70)
    print("✅ Compilación exitosa!")
    print(f"El paquete portable se encuentra en:\n   {dist_dir / 'ReconcileAI'}")
    print("Para usarlo en cualquier computadora:")
    print("1. Copia la carpeta 'dist/ReconcileAI' a una memoria USB o a otra PC.")
    print("2. Haz doble clic en 'ReconcileAI.exe'.")
    print("=" * 70)

if __name__ == "__main__":
    build()
