import sys
from PyInstaller.__main__ import run

sys.argv = ["--noconfirm", "--clean", "build_exe.spec"]
run(sys.argv)
