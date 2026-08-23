import sys

print("python", sys.executable, sys.version)
try:
    import tkinter  # noqa: F401
    print("tkinter OK")
except Exception as e:
    print("tkinter FAIL:", type(e).__name__, e)
try:
    import PyInstaller  # noqa: F401
    print("pyi OK", PyInstaller.__version__)
except Exception as e:
    print("pyi FAIL:", type(e).__name__, e)
try:
    import pyi  # noqa
    print("pyi module ok")
except Exception as e:
    print("pyi module FAIL:", type(e).__name__, e)
