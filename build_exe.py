"""
Build script to create standalone executable for Universal Email Sender
This will create a single .exe file with all dependencies included
"""
import os
import sys
import subprocess
import shutil
def check_pyinstaller():
    """Check if PyInstaller is installed"""
    try:
        import PyInstaller
        print("[OK] PyInstaller is installed")
        return True
    except ImportError:
        print("[!] PyInstaller is not installed")
        return False
def install_pyinstaller():
    """Install PyInstaller"""
    print("\nInstalling PyInstaller...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("[OK] PyInstaller installed successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"[!] Failed to install PyInstaller: {e}")
        return False
def clean_build_folders():
    """Clean up old build folders"""
    folders_to_remove = ['build', 'dist', '__pycache__']
    for folder in folders_to_remove:
        if os.path.exists(folder):
            print(f"Cleaning {folder}...")
            shutil.rmtree(folder, ignore_errors=True)
    if os.path.exists('MailMergeSender.spec'):
        os.remove('MailMergeSender.spec')
def verify_runtime_hook():
    """Verify the runtime hook file exists"""
    if os.path.exists('pyi_rth_win32com.py'):
        print("[OK] Runtime hook found: pyi_rth_win32com.py")
        return True
    else:
        print("[!] Runtime hook not found: pyi_rth_win32com.py")
        print("  This file is required for Outlook integration in EXE mode")
        return False
def build_executable():
    """Build the executable using PyInstaller"""
    print("\n" + "="*60)
    print("BUILDING UNIVERSAL EMAIL SENDER APPLICATION")
    print("="*60 + "\n")
    if not verify_runtime_hook():
        return False
    cmd = [
        'pyinstaller',
        '--onedir',                     
        '--windowed',                   
        '--name=MailMergeSender',  
        '--icon=NONE',                  
        '--runtime-hook=pyi_rth_win32com.py',
        '--hidden-import=theme',
        '--hidden-import=loading_screen',
        '--hidden-import=mail_merge_sender',
        '--hidden-import=PyQt5',
        '--hidden-import=PyQt5.QtCore',
        '--hidden-import=PyQt5.QtGui',
        '--hidden-import=PyQt5.QtWidgets',
        '--hidden-import=win32com',
        '--hidden-import=win32com.client',
        '--hidden-import=pythoncom',
        '--hidden-import=pywintypes',
        '--hidden-import=pandas',
        '--hidden-import=openpyxl',
        '--hidden-import=docx',
        '--collect-all=pywin32',
        '--copy-metadata=pywin32',
        '--exclude-module=matplotlib',
        '--exclude-module=numpy.distutils',
        '--exclude-module=PIL',
        '--exclude-module=scipy',
        '--exclude-module=IPython',
        'main.py'
    ]
    print("Running PyInstaller with the following command:")
    print(" ".join(cmd))
    print("\n" + "="*60 + "\n")
    try:
        subprocess.check_call(cmd)
        print("\n" + "="*60)
        print("[SUCCESS] BUILD SUCCESSFUL!")
        print("="*60)
        return True
    except subprocess.CalledProcessError as e:
        print("\n" + "="*60)
        print(f"[ERROR] BUILD FAILED: {e}")
        print("="*60)
        return False
def verify_build():
    """Verify the executable and folder were created"""
    exe_path = os.path.join('dist', 'MailMergeSender', 'MailMergeSender.exe')
    folder_path = os.path.join('dist', 'MailMergeSender')
    if os.path.exists(exe_path) and os.path.isdir(folder_path):
        total_size = 0
        for dirpath, dirnames, filenames in os.walk(folder_path):
            for filename in filenames:
                filepath = os.path.join(dirpath, filename)
                total_size += os.path.getsize(filepath)
        size_mb = total_size / (1024 * 1024)
        file_count = len([f for _, _, files in os.walk(folder_path) for f in files])
        print(f"\n[OK] Application folder created successfully!")
        print(f"  Location: {os.path.abspath(folder_path)}")
        print(f"  Executable: {os.path.abspath(exe_path)}")
        print(f"  Total size: {size_mb:.2f} MB ({file_count} files)")
        return True
    else:
        print(f"\n[!] Application folder not found at: {folder_path}")
        return False
def main():
    print("="*60)
    print("Universal Email Sender - Executable Builder")
    print("="*60 + "\n")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}\n")
    required_files = ['main.py', 'mail_merge_sender.py', 'theme.py', 'loading_screen.py']
    missing_files = [f for f in required_files if not os.path.exists(f)]
    if missing_files:
        print(f"[!] Missing required files: {', '.join(missing_files)}")
        return False
    print("[OK] All required files found\n")
    if not check_pyinstaller():
        if not install_pyinstaller():
            print("\nPlease install PyInstaller manually: pip install pyinstaller")
            return False
    print("\nCleaning old build files...")
    clean_build_folders()
    print("[OK] Cleanup complete\n")
    if not build_executable():
        return False
    if not verify_build():
        return False
    print("\n" + "="*60)
    print("BUILD PROCESS COMPLETE")
    print("="*60)
    print("\nYou can now run the application from:")
    print(f"  dist\\MailMergeSender\\MailMergeSender.exe")
    print("\nThe application folder includes:")
    print("  • Main executable (MailMergeSender.exe)")
    print("  • All Python code and dependencies")
    print("  • PyQt5 GUI framework")
    print("  • pywin32 for Outlook integration")
    print("  • pandas, openpyxl for data import")
    print("  • python-docx for Word document support")
    print("\nTo distribute:")
    print("  Copy the entire 'dist\\MailMergeSender' folder")
    print("  to any Windows PC and run the .exe - no installation needed!")
    print("="*60 + "\n")
    return True
if __name__ == "__main__":
    try:
        success = main()
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\nBuild cancelled by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n\nUnexpected error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
