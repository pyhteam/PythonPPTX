# setup.py
import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine-tuning.
build_exe_options = {"packages": [], "excludes": []}

# Base setting for Windows
base = None
if sys.platform == "win32":
    base = "Win32GUI"  # Use "Win32GUI" if your script is a GUI application

setup(
    name = "export_pptx_library",
    version = "0.1",
    description = "Library support export pptx",
    options = {"build_exe": build_exe_options},
    executables = [Executable("buil_to_exe.py", base=base)],
)
