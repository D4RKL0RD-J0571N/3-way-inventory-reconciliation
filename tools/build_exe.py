"""
Automated PyInstaller build script for ReconcileAI.
Packages the entire application into a portable, standalone executable directory (dist/ReconcileAI).
No Python installation required on the destination machine.
"""
import sys
import os
import shutil
from pathlib import Path

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import PyInstaller.__main__
from PyInstaller.utils.hooks import copy_metadata, collect_data_files, collect_submodules

BASE_DIR = Path(__file__).resolve().parent.parent

def build():
    print("=" * 70)
    print("  CREANDO EJECUTABLE PORTABLE DE RECONCILEAI (PyInstaller)")
    print("=" * 70)

    # Clean previous build artifacts
    build_dir = BASE_DIR / "build"
    dist_dir = BASE_DIR / "dist"
    app_dist_dir = dist_dir / "ReconcileAI"

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

    # Submodules for hidden imports
    hidden_imports = [
        "streamlit",
        "openpyxl",
        "pandas",
        "altair",
        "pyarrow",
        "tornado",
        *collect_submodules("streamlit")
    ]

    # Convert datas to PyInstaller CLI arguments
    data_args = []
    for src, dst in datas:
        data_args.extend(["--add-data", f"{src}{os.pathsep}{dst}"])

    hidden_args = []
    for mod in set(hidden_imports):
        hidden_args.extend(["--hidden-import", mod])

    pyinstaller_args = [
        str(BASE_DIR / "run_standalone.py"),
        "--name=ReconcileAI",
        "--onedir",
        "--noconfirm",
        "--clean",
        *hidden_args,
        *data_args
    ]

    print("Ejecutando PyInstaller (esto puede tardar 1-2 minutos)...")
    PyInstaller.__main__.run(pyinstaller_args)

    # Post-build: Create local data folder & helper launcher inside dist/ReconcileAI
    target_data_dir = app_dist_dir / "data"
    target_data_dir.mkdir(parents=True, exist_ok=True)
    
    readme_data = target_data_dir / "LEEME.txt"
    readme_data.write_text(
        "Coloca aqui las carpetas de tus importaciones o contenedores (por ejemplo CONT-001, CONT-002).\n"
        "La aplicacion detectara automaticamente los archivos de Control de Calidad y Plantilla de Cierre.\n"
        "Tambien puedes explorar cualquier carpeta de tu disco duro desde la aplicacion web.",
        encoding="utf-8"
    )

    launcher_bat = app_dist_dir / "INICIAR_RECONCILEAI.bat"
    launcher_bat.write_text(
        "@echo off\r\n"
        "title ReconcileAI - Copiloto 3-Way\r\n"
        "cd /d \"%~dp0\"\r\n"
        "echo Iniciando ReconcileAI...\r\n"
        "start \"\" \"ReconcileAI.exe\"\r\n"
        "exit /b 0\r\n",
        encoding="utf-8"
    )

    readme_root = app_dist_dir / "LEEME_PORTABLE.txt"
    readme_root.write_text(
        "======================================================================\n"
        "  ReconcileAI - Version Portable Independiente (Windows 64-bit)\n"
        "======================================================================\n\n"
        "COMO USAR ESTA VERSION:\n"
        "1. No necesitas tener Python instalado en tu computadora.\n"
        "2. No necesitas permisos de Administrador de Windows.\n"
        "3. Simplemente haz doble clic en 'ReconcileAI.exe' (o en 'INICIAR_RECONCILEAI.bat').\n"
        "4. Tu navegador web se abrira inmediatamente con la aplicacion lista.\n\n"
        "DONDE COLOCAR TUS ARCHIVOS EXCEL:\n"
        "- Puedes arrastrarlos directamente en la pestana 'Subir Archivos'.\n"
        "- O copiar tu carpeta (ej. CONT-001) dentro de la carpeta 'data/'.\n"
        "- O escribir la ruta completa en el explorador de carpetas locales.\n",
        encoding="utf-8"
    )

    # Create ZIP archive for standalone portable distribution
    zip_path = dist_dir / "ReconcileAI_Portable"
    print("\nGenerando archivo comprimido ZIP portable...")
    shutil.make_archive(str(zip_path), "zip", dist_dir, "ReconcileAI")

    final_zip = dist_dir / "ReconcileAI_Portable.zip"
    print("\n" + "=" * 70)
    print("COMPILACION Y EMPAQUETADO EXITOSO!")
    print(f"1. Carpeta Portable: {app_dist_dir}")
    print(f"2. Archivo ZIP Portable: {final_zip} ({final_zip.stat().st_size / (1024*1024):.1f} MB)")
    print("=" * 70)

if __name__ == "__main__":
    build()
