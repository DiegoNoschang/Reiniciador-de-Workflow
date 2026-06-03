# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec — gera um .exe único do Reiniciador de Workflow - iiLex (RPA).

Como usar:
    pip install pyinstaller
    pyinstaller iilex.spec

O executável final fica em:
    dist/Reiniciador de Workflow - iiLex (RPA).exe
"""

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# ============================================================
# Hidden imports — bibliotecas com imports dinâmicos
# ============================================================
# Selenium e webdriver_manager usam imports dinâmicos em vários
# pontos (factories, plug-ins). Sem isso o .exe quebra com
# "No module named 'selenium.webdriver.chrome.webdriver'".
hidden_selenium = collect_submodules('selenium')
hidden_wdm = collect_submodules('webdriver_manager')
hidden_openpyxl = collect_submodules('openpyxl')

hidden_imports = [
    # pywin32 (DPAPI para criptografar senhas)
    'win32crypt',
    'win32timezone',
    'pywintypes',
    'pythoncom',
] + hidden_selenium + hidden_wdm + hidden_openpyxl

# ============================================================
# Data files — arquivos de dados que precisam ir junto
# ============================================================
extra_datas = [
    # Logos e ícones (header, footer, window icon)
    ('assets/ramos.ico', 'assets'),
    ('assets/R_logo.png', 'assets'),
    ('assets/logo_ramos.png', 'assets'),
]

# selenium e webdriver_manager têm arquivos auxiliares (templates,
# JSONs de configuração, etc.) que precisam estar bundleados.
extra_datas += collect_data_files('selenium')
extra_datas += collect_data_files('webdriver_manager')


a = Analysis(
    ['interface_iilex_qt.py'],
    pathex=[],
    binaries=[],
    datas=extra_datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Reduz o tamanho do .exe — tira coisas que não usamos
        'tkinter',
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'PIL',
        'PyQt5',
        'PySide6',
        'PySide2',
        'IPython',
        'jupyter',
        'notebook',
    ],
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
    name='Reiniciador de Workflow - iiLex (RPA)',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,           # GUI app, sem console
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/ramos.ico',  # ícone do .exe (aparece no Explorer)
)
