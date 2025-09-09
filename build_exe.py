#!/usr/bin/env python3
"""
Script de compilation pour créer un exécutable .EXE de l'application excel_to_word_app.py
"""

import os
import sys
import subprocess
import shutil

def install_pyinstaller():
    """Installe PyInstaller si nécessaire"""
    try:
        import PyInstaller
        print("OK PyInstaller est deja installe")
    except ImportError:
        print("Installation de PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("OK PyInstaller installe avec succes")

def create_spec_file():
    """Crée le fichier .spec pour PyInstaller"""
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['excel_to_word_app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('logo-Toyota-Solo.ico', '.'),
        ('LogoTMH.png', '.'),
        ('Datas', 'Datas'),
    ],
    hiddenimports=[
        'pandas',
        'openpyxl',
        'docx',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        'numpy',
        'tkinter',
        'tkinter.filedialog',
        'tkinter.messagebox',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Generator_Formation',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='logo-Toyota-Solo.ico'
)
'''
    
    with open('excel_to_word_app.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    print("OK Fichier .spec cree")

def build_exe():
    """Compile l'application en .EXE"""
    print("Debut de la compilation...")
    
    # Utiliser le fichier .spec pour la compilation
    result = subprocess.run([
        'pyinstaller',
        '--clean',
        'excel_to_word_app.spec'
    ], capture_output=True, text=True)
    
    if result.returncode == 0:
        print("OK Compilation reussie!")
        print(f"L'exécutable se trouve dans: dist/Generator_Formation.exe")
        return True
    else:
        print("ERREUR lors de la compilation:")
        print(result.stderr)
        return False

def cleanup():
    """Nettoie les fichiers temporaires"""
    print("Nettoyage des fichiers temporaires...")
    
    # Supprimer le dossier build s'il existe
    try:
        if os.path.exists('build'):
            shutil.rmtree('build')
            print("OK Dossier 'build' supprime")
    except PermissionError:
        print("ATTENTION: Impossible de supprimer le dossier 'build' (fichiers en cours d'utilisation)")
    except Exception as e:
        print(f"ATTENTION: Erreur lors de la suppression du dossier 'build': {e}")
    
    # Supprimer le fichier .spec
    try:
        if os.path.exists('excel_to_word_app.spec'):
            os.remove('excel_to_word_app.spec')
            print("OK Fichier .spec supprime")
    except Exception as e:
        print(f"ATTENTION: Erreur lors de la suppression du fichier .spec: {e}")

def main():
    """Fonction principale"""
    print("=== Compilation de l'application Generator Formation ===")
    print()
    
    # Vérifier que le fichier principal existe
    if not os.path.exists('excel_to_word_app.py'):
        print("ERREUR: Le fichier excel_to_word_app.py n'existe pas!")
        return
    
    # Vérifier que les ressources existent
    required_files = ['logo-Toyota-Solo.ico', 'LogoTMH.png', 'Datas']
    missing_files = []
    
    for file in required_files:
        if not os.path.exists(file):
            missing_files.append(file)
    
    if missing_files:
        print(f"ERREUR: Fichiers manquants: {', '.join(missing_files)}")
        print("Assurez-vous que tous les fichiers requis sont présents dans le dossier.")
        return
    
    try:
        # Installer PyInstaller
        install_pyinstaller()
        print("")
        
        # Créer le fichier .spec
        create_spec_file()
        print("")
        
        # Compiler l'application
        if build_exe():
            print("")
            print("=== SUCCES ===")
            print("Votre application a ete compilee avec succes!")
            print("L'exécutable se trouve dans: dist/Generator_Formation.exe")
            print("")
            print("Pour distribuer l'application:")
            print("1. Copiez le fichier dist/Generator_Formation.exe")
            print("2. Copiez le dossier Datas/ (s'il n'est pas inclus)")
            print("3. L'application peut être exécutée sur n'importe quel PC Windows")
        else:
            print("ERREUR: La compilation a echoue")
            
    except Exception as e:
        print(f"ERREUR: {e}")
    
    finally:
        # Nettoyer les fichiers temporaires
        print()
        cleanup()

if __name__ == "__main__":
    main()

