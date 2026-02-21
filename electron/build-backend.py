import subprocess
import sys
import os
import shutil
from pathlib import Path

ROOT = Path(__file__).parent.parent
BACKEND = ROOT / "backend"
DIST = ROOT / "electron" / "backend-dist"
WORK = ROOT / "electron" / "build-temp"

# Clean previous builds
if DIST.exists():
    shutil.rmtree(DIST)
if WORK.exists():
    shutil.rmtree(WORK)

print(f"Python version: {sys.version}")
print(f"Building from: {BACKEND}")
print(f"Output to: {DIST}")

# Run version check first
result = subprocess.run([sys.executable, str(BACKEND / "python_version_check.py")])
if result.returncode != 0:
    sys.exit(1)

cmd = [
    sys.executable, "-m", "PyInstaller",
    "--onedir",
    "--name", "mailscraper-backend",
    "--distpath", str(DIST),
    "--workpath", str(WORK),
    "--specpath", str(ROOT / "electron"),
    "--clean",
    "--noconfirm",
    # Uvicorn hidden imports
    "--hidden-import=uvicorn.logging",
    "--hidden-import=uvicorn.loops",
    "--hidden-import=uvicorn.loops.auto",
    "--hidden-import=uvicorn.protocols",
    "--hidden-import=uvicorn.protocols.http",
    "--hidden-import=uvicorn.protocols.http.auto",
    "--hidden-import=uvicorn.protocols.websockets",
    "--hidden-import=uvicorn.protocols.websockets.auto",
    "--hidden-import=uvicorn.lifespan",
    "--hidden-import=uvicorn.lifespan.on",
    # FastAPI + SQLAlchemy
    "--hidden-import=fastapi",
    "--hidden-import=sqlalchemy.dialects.sqlite",
    "--hidden-import=sqlalchemy.orm",
    # Email + parsing
    "--hidden-import=imaplib",
    "--hidden-import=email",
    "--hidden-import=email.header",
    "--hidden-import=email.utils",
    "--hidden-import=pdfplumber",
    "--hidden-import=docx",
    # Scheduler + auth
    "--hidden-import=apscheduler",
    "--hidden-import=apscheduler.schedulers.asyncio",
    "--hidden-import=apscheduler.triggers.interval",
    "--hidden-import=msal",
    "--hidden-import=requests",
    # Data
    "--hidden-import=pandas",
    "--hidden-import=openpyxl",
    # Collect all data files
    "--collect-all=pdfplumber",
    "--collect-all=docx",
    "--collect-all=uvicorn",
    "--collect-all=fastapi",
    "--collect-all=msal",
    "--collect-data=apscheduler",
    str(BACKEND / "main.py"),
]

print("Running PyInstaller...")
result = subprocess.run(cmd, cwd=str(BACKEND))

if result.returncode != 0:
    print("\nPyInstaller FAILED. Check errors above.")
    sys.exit(1)

# Verify output
binary = DIST / "mailscraper-backend" / ("mailscraper-backend.exe" if sys.platform == "win32" else "mailscraper-backend")
if not binary.exists():
    print(f"ERROR: Expected binary not found at {binary}")
    sys.exit(1)

# Copy .env.template into the dist folder
template = BACKEND / ".env.template"
if template.exists():
    shutil.copy(template, DIST / "mailscraper-backend" / ".env.template")
    print("Copied .env.template to dist")

print(f"\nBuild successful!")
print(f"Binary: {binary}")
print(f"Size: {binary.stat().st_size / 1024 / 1024:.1f} MB")
