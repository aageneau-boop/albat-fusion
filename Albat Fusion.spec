# -*- mode: python ; coding: utf-8 -*-
#
# Fichier de configuration PyInstaller pour Albat Fusion.
#
# Utilisation (depuis une invite de commandes Windows, dans le
# dossier contenant ce fichier ET "Albat_Fusion_Mac.py") :
#
#     pyinstaller "Albat Fusion.spec"
#
# Le résultat se trouve dans dist/Albat Fusion/ — c'est CE DOSSIER
# ENTIER qu'il faut distribuer (mode --onedir), pas un fichier
# .exe isolé : resource_path() a besoin des dossiers assets/,
# profils/ et "dossiers pour rapport/" à côté de l'exécutable pour
# fonctionner, et le mode --onefile les aurait rendus inaccessibles
# d'une session à l'autre (dossier temporaire recréé à chaque
# lancement).

import sys
from PyInstaller.utils.hooks import collect_submodules

block_cipher = None

# Sécurité supplémentaire : les 5 modules d'Albat Fusion sont
# importés à l'intérieur de Suite.__init__ (imports différés pour
# accélérer le démarrage), pas en haut du fichier. L'analyse
# statique de PyInstaller les détecte normalement quand même, mais
# on les liste explicitement au cas où.
hidden_imports = (
    collect_submodules('docx') +
    collect_submodules('openpyxl') +
    collect_submodules('astral') +
    collect_submodules('matplotlib') +
    [
        'modules',
        'modules.utils',
        'modules.Albat_Correlations_Fusion_Mac',
        'modules.Albat_Graph_Fusion_Mac',
        'modules.Albat_Bridage_Fusion_Mac',
        'modules.Albat_Scenar_Mac',
        'modules.Albat_Rapport_Mac',
        'PIL',
        'PIL.Image',
        'PIL._tkinter_finder',
        'pandas',
        'zoneinfo',
    ]
)

# Dossiers/fichiers copiés tels quels dans le build final, à côté
# de l'exécutable. Doivent exister (même vides, avec juste un
# fichier .gitkeep dedans par exemple) avant de lancer pyinstaller.
datas = [
    ('assets', 'assets'),
    ('profils', 'profils'),
    ('dossiers pour rapport', 'dossiers pour rapport'),
]

a = Analysis(
    ['Albat_Fusion_Mac.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
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
    [],
    exclude_binaries=True,
    name='Albat Fusion',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,          # pas de fenêtre console noire au lancement
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/icone_albat.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Albat Fusion',
)
