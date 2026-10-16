# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('assets', 'assets/'), ('data', 'data/'), ('scripts', 'scripts/'), ('modulos', 'modulos/'), ('skills', 'skills/'), ('versao.json', '.')]
binaries = []
hiddenimports = [
    'plyer', 'win10toast', 'pythoncom', 'win32com', 'win32clipboard',
    'openpyxl', 'pandas', 'PIL', 'reportlab',
    'selenium', 'selenium.webdriver', 'selenium.webdriver.chrome',
    'selenium.webdriver.chrome.options', 'selenium.webdriver.chrome.service',
    'selenium.webdriver.chrome.webdriver',
    'modulos.compilador_plataforma', 'modulos.mudanca_recurso', 'modulo_plataforma', 'modulo_consultar_migo',
    'modulo_saida_osme',
    'modulo_contagem_carro', 'modulo_fichas', 'modulo_dashboard',
    'modulo_conferencia', 'modulo_bandeirada', 'modulo_etiquetas',
    'modulo_extrair_OSME', 'modulo_extrator', 'modulo_Sap', 'modulo_chat',
    'Modulo_ZMCSE003', 'modulo_transferenciaosme', 'modulo_mudanca_recurso',
    'modulo_atualizar_osme', 'modulo_contagem_deposito',
    'utils_os', 'utils_skills', 'utils_pendencias_planilha', 'utils_metrics', 'utils_workforce', 'sidebar_navegacao', 'yaml',
    'pystray', 'utils_tray', 'utils_updater', 'utils_som', 'utils_backup_diario', 'utils_spotlight'
]

tmp_ret = collect_all('reportlab')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret2 = collect_all('selenium')
datas += tmp_ret2[0]; binaries += tmp_ret2[1]; hiddenimports += tmp_ret2[2]
tmp_ret3 = collect_all('customtkinter')
datas += tmp_ret3[0]; binaries += tmp_ret3[1]; hiddenimports += tmp_ret3[2]
tmp_ret4 = collect_all('pystray')
datas += tmp_ret4[0]; binaries += tmp_ret4[1]; hiddenimports += tmp_ret4[2]

a = Analysis(
    ['src\\conecta_hub2.py'],
    pathex=['src', '.'],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='N.E.X.U.S',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['assets\\Nexus.ico']
)
