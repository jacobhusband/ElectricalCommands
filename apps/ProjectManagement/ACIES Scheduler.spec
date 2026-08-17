# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('VERSION', '.'), ('index.html', '.'), ('styles.css', '.'), ('script.js', '.'), ('.env', '.'), ('assets\\acies.png', 'assets'), ('assets\\lighting', 'assets\\lighting'), ('scripts\\merge_pdfs.py', 'scripts'), ('scripts\\strip_pdf_layers.py', 'scripts'), ('scripts\\detect_pdf_size.py', 'scripts'), ('scripts\\PlotDWGs.ps1', 'scripts'), ('scripts\\ManageLayersDWGs.ps1', 'scripts'), ('scripts\\ListDwgXrefs.ps1', 'scripts'),('scripts\\removeXREFPaths.ps1', 'scripts'), ('scripts\\StripRefPaths.dll', 'scripts'), ('templates', 'templates'), ('CircuitBreakerAI\\ElectricalPanels\\Template.xlsx', 'CircuitBreakerAI\\ElectricalPanels'), ('WireSizerApplication\\\\dist', 'WireSizerApplication\\\\dist'), ('project-pages-editor\\\\dist', 'project-pages-editor\\\\dist')],
    hiddenimports=['pillow_heif', '_pillow_heif'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ACIES Scheduler',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['assets\\acies.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ACIES Scheduler',
)
