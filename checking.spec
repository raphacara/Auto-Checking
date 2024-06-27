# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['checking.py'],  # Assurez-vous que c'est le bon fichier
    pathex=['.'],  # Chemin vers le r�pertoire de votre projet
    binaries=[],
    datas=[
        ('engine.ico', '.'),
        ('engine.png', '.'),
        ('intro1.png', '.'),
        ('button_1_data.txt', '.'),  # Inclure tous les fichiers de donn�es n�cessaires
        ('button_2_data.txt', '.'),
        ('button_3_data.txt', '.'),
        ('button_4_data.txt', '.'),
        ('button_5_data.txt', '.'),
        ('button_6_data.txt', '.'),
        ('button_7_data.txt', '.'),
    ],
    hiddenimports=[],  # Ajoutez ici les modules Python import�s dynamiquement
    hookspath=[],  # Chemin vers les hooks personnalis�s si n�cessaire
    hooksconfig={},
    runtime_hooks=[],  # Ajoutez des hooks d'ex�cution si n�cessaire
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='checking',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Changez � True si vous souhaitez une fen�tre console
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='engine.ico'  # Assurez-vous que le chemin vers l'ic�ne est correct
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Checking'
)

app = BUNDLE(
    coll,
    name='checking.app',  # Nom de l'application
    icon='engine.ico',  # Chemin vers l'ic�ne
    bundle_identifier=None,
    info_plist=None,
    runtime_tmpdir=None,
    options={'bundle_files': 1}  # Option pour cr�er un ex�cutable unique
)
