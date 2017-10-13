from cx_Freeze import setup, Executable
import sys

base = None
if sys.platform == 'win32':
    base = 'Win32GUI'

setup(
    name = "Траты по бюджету",
    version = "0.1",
    description = "Budget",
    executables = [Executable(script="hello.py", base=base)]
)
