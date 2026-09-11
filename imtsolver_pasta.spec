# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = []
import os
import pulp
# o CBC que vem dentro do PuLP: achado pelo pacote instalado, e nao por
# um caminho fixo desta maquina
_cbc = os.path.join(os.path.dirname(pulp.__file__), 'solverdir', 'cbc', 'win', 'i64', 'cbc.exe')
binaries = [(_cbc, '.')]
hiddenimports = []
tmp_ret = collect_all('pyscipopt')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('highspy')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['byesolver.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'pandas'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

# Pasta em vez de arquivo unico. No modo onefile o PyInstaller
# descompacta os 79 MB para uma pasta temporaria a cada execucao, o que
# custava 1,42 s antes da primeira instrucao rodar. Em pasta, as DLLs ja
# estao no disco e o processo comeca direto.
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='imtsolver',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
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
    name='imtsolver',
)
