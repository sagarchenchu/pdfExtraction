# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for PDF Table Extractor.

Build the standalone executable with:
    pyinstaller pdf_extractor.spec

The output will be in the dist/ folder.
"""

block_cipher = None

a = Analysis(
    ["pdf_extractor.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        "pdfplumber",
        "pdfminer",
        "pdfminer.high_level",
        "pdfminer.layout",
        "pdfminer.converter",
        "pdfminer.pdfpage",
        "pdfminer.pdfinterp",
        "pdfminer.image",
        "openpyxl",
        "openpyxl.styles",
        "openpyxl.utils",
        "docx",
        "docx.oxml",
        "docx.oxml.ns",
        "concurrent.futures",
        "PIL",
        "Wand",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="PDFTableExtractor",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # No console window – GUI only
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon="icon.ico",      # Uncomment and provide an .ico file to set a custom icon
)
