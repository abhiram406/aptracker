# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main_gui_edits.py'],
    pathex=[],
    binaries=[],
    datas=[('BEL-Logo-PNG.png', '.'), ('kas-logo.png', '.'), ('template_SFB1_RF.docx', '.'), ('template_SFB1_BITE.docx', '.'), ('template_SFB2_RF.docx', '.'), ('template_SFB2_BITE.docx', '.'), ('template_SFB3_RF.docx', '.'), ('template_SFB3_BITE.docx', '.'), ('template_BITE.docx', '.'), ('template_RF.docx', '.'), ('template_HMR_SFB_Low_RF.docx', '.'), ('template_HMR_SFB_Low_BITE.docx', '.'), ('template_HMR_SFB_High_RF.docx', '.'), ('template_HMR_SFB_High_BITE.docx', '.'), ('star-32.ico', '.')],
    hiddenimports=[],
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
    name='Amp & Phase Tracking',
    icon='star-32.ico',
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
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='main_gui_edits',
)
