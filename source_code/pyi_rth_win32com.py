"""
PyInstaller runtime hook for win32com COM initialization.
This ensures COM is properly initialized when running as a frozen executable.
CRITICAL: This runs BEFORE any application code, setting up win32com for proper
Outlook integration in EXE mode. This is the single point of COM initialization.
"""
import sys
import os
if hasattr(sys, 'frozen'):
    try:
        import pythoncom
        pythoncom.CoInitialize()
        print("Runtime hook: COM initialized with STA threading model")
        try:
            import win32com.client.gencache
            win32com.client.gencache.is_readonly = True
            win32com.client.gencache.__gen_path__ = ""
            print("Runtime hook: Configured gencache for frozen mode")
        except Exception as e:
            print(f"Runtime hook: gencache config note - {e}")
        try:
            import win32com.client
            print("Runtime hook: win32com.client loaded for dynamic dispatch")
        except Exception as e:
            print(f"Runtime hook: win32com.client config note - {e}")
        print("SUCCESS: Runtime hook - COM fully initialized for Outlook automation")
    except Exception as e:
        print(f"ERROR: Runtime hook - {e}")
        import traceback
        traceback.print_exc()
