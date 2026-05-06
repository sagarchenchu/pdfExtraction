"""
build_exe.py – Helper script to install dependencies and build the
standalone executable with PyInstaller.

Usage (run from the repository root):
    python build_exe.py

The finished executable will be placed in:
    dist/PDFTableExtractor.exe   (Windows)
    dist/PDFTableExtractor       (macOS / Linux)
"""

import subprocess
import sys
import os

ROOT = os.path.dirname(os.path.abspath(__file__))


def run(cmd: list[str]) -> None:
    print(f"\n>>> {' '.join(cmd)}\n")
    subprocess.run(cmd, check=True)


def main() -> None:
    # 1. Install runtime dependencies (everything except pyinstaller itself)
    deps = ["pdfplumber", "openpyxl", "python-docx"]
    run([sys.executable, "-m", "pip", "install", "--upgrade"] + deps)

    # 2. Install PyInstaller
    run([sys.executable, "-m", "pip", "install", "--upgrade", "pyinstaller"])

    # 3. Build
    spec_file = os.path.join(ROOT, "pdf_extractor.spec")
    run([sys.executable, "-m", "PyInstaller", spec_file, "--clean"])

    print("\n✅  Build finished. Executable is in the dist/ folder.\n")


if __name__ == "__main__":
    main()
