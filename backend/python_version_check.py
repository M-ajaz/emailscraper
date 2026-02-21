import sys
if sys.version_info >= (3, 12):
    print(f"ERROR: Python {sys.version} is not supported.")
    print("Please install Python 3.11 from https://www.python.org/downloads/release/python-3119/")
    print("During install, CHECK the 'Add Python to PATH' checkbox.")
    sys.exit(1)
print(f"Python {sys.version} — OK")
