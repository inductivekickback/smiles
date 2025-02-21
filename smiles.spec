# -*- mode: python ; coding: utf-8 -*-
import sys
from pathlib import Path

sys.path.append(str(Path('.').resolve()))
import smiles

a = Analysis(
    ['smiles.py', 'pdf_writer.py'],
    pathex=[],
    binaries=[],
    datas=[('artifacts/20250218_mileage.pdf', 'artifacts'),
                ('artifacts/20250218_additional_mileage.pdf', 'artifacts'),
                ('artifacts/20250218_distances.pdf', 'artifacts'),
                ('artifacts/support.png', 'artifacts'),
                ('installer_bg.png', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=2,
    upx=True,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Smiles',
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
    strip=True,
    upx=True,
    upx_exclude=[],
    name='Smiles',
    runtime_tmpdir=None,
    noarchive=False,
)

plist = {
    'NSPrincipalClass': 'NSApplication',
    'NSAppleScriptEnabled': False,
    'CFBundleName': 'Smiles',
    'CFBundleIdentifier': 'io.foolsday.smiles',
    'CFBundleVersion': f'{smiles.__version__}',
    'CFBundleShortVersionString': f'{smiles.__version__}',
    'CFBundleDocumentTypes': [
        {
            'CFBundleTypeName': 'RLM File',
            'CFBundleTypeRole': 'Viewer',
            'LSHandlerContentType': 'io.foolsday.rlm',
            'LSHandlerRank': 'Owner',
            'CFBundleTypeExtensions': ['rlm'],
        }
    ],
    'LSItemContentTypes': [
        'io.foolsday.rlm'
    ],
    'UTExportedTypeDeclarations': [
        {
            'UTTypeTagSpecification': {
                'public.filename-extension': ['rlm']
            },
            'UTTypeIdentifier': 'io.foolsday.rlm',
            'UTTypeDescription': 'RLM File',
            'UTTypeConformsTo': ['public.data']
        }
    ]
}

app = BUNDLE(
    coll,
    name='Smiles.app',
    icon='mentor.icns',
    bundle_identifier=None,
    info_plist=plist,
)
