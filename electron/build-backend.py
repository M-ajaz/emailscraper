import subprocess
import sys
import os
import shutil
from pathlib import Path

ROOT = Path(__file__).parent.parent
BACKEND = ROOT / "backend"
DIST = ROOT / "electron" / "backend-dist"

print("Building Python backend with PyInstaller...")

cmd = [
    sys.executable, "-m", "PyInstaller",
    "--onedir",
    "--name", "mailscraper-backend",
    "--distpath", str(DIST),
    "--workpath", str(ROOT / "electron" / "build-temp"),
    "--specpath", str(ROOT / "electron"),
    "--hidden-import", "uvicorn.logging",
    "--hidden-import", "uvicorn.loops",
    "--hidden-import", "uvicorn.loops.auto",
    "--hidden-import", "uvicorn.protocols",
    "--hidden-import", "uvicorn.protocols.http",
    "--hidden-import", "uvicorn.protocols.http.auto",
    "--hidden-import", "uvicorn.protocols.websockets",
    "--hidden-import", "uvicorn.protocols.websockets.auto",
    "--hidden-import", "uvicorn.lifespan",
    "--hidden-import", "uvicorn.lifespan.on",
    "--hidden-import", "sqlalchemy.dialects.sqlite",
    "--hidden-import", "pdfplumber",
    "--hidden-import", "docx",
    "--hidden-import", "apscheduler",
    "--hidden-import", "msal",
    "--hidden-import", "imaplib",
    "--hidden-import", "email",
    "--collect-all", "pdfplumber",
    "--collect-all", "docx",
    str(BACKEND / "main.py"),
]

result = subprocess.run(cmd, cwd=str(BACKEND))
if result.returncode != 0:
    print("PyInstaller failed!")
    sys.exit(1)

print("Backend bundled successfully to:", DIST)
