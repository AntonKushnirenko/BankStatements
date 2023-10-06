from kivy_deps import sdl2, glew

# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['BankStatements.py'],
    pathex=[],
    binaries=[],
    datas=[('bankstatements.kv', '.'), ('select_a_file_down.png', '.'), ('select_a_file_down_blank.png', '.'), ('select_a_file_normal.png', '.'), ('select_a_file_normal_blank.png', '.'), ('service_account.json', '.'), ('upload_to_google_disabled.png', '.'), ('upload_to_google_down.png', '.'), ('upload_to_google_normal.png', '.'), ('artel_icon.png', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=True,
)

a.datas += [('files/select_a_file_down.png', 'C:\\Users\\anton\\Desktop\\bankstatements\\select_a_file_down.png', 'DATA')]
a.datas += [('files/select_a_file_down_blank.png', 'C:\\Users\\anton\\Desktop\\bankstatements\\select_a_file_down_blank.png', 'DATA')]
a.datas += [('files/select_a_file_normal.png', 'C:\\Users\\anton\\Desktop\\bankstatements\\select_a_file_normal.png', 'DATA')]
a.datas += [('files/select_a_file_normal_blank.png', 'C:\\Users\\anton\\Desktop\\bankstatements\\select_a_file_normal_blank.png', 'DATA')]
a.datas += [('files/upload_to_google_disabled.png', 'C:\\Users\\anton\\Desktop\\bankstatements\\upload_to_google_disabled.png', 'DATA')]
a.datas += [('files/upload_to_google_down.png', 'C:\\Users\\anton\\Desktop\\bankstatements\\upload_to_google_down.png', 'DATA')]
a.datas += [('files/upload_to_google_normal.png', 'C:\\Users\\anton\\Desktop\\bankstatements\\upload_to_google_normal.png', 'DATA')]
a.datas += [('files/artel_icon.png', 'C:\\Users\\anton\\Desktop\\bankstatements\\artel_icon.png', 'DATA')]
a.datas += [('files/service_account.json', 'C:\\Users\\anton\\Desktop\\bankstatements\\service_account.json', 'DATA')]
a.datas += [('files/bankstatements.kv', 'C:\\Users\\anton\\Desktop\\bankstatements\\bankstatements.kv', 'DATA')]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz, Tree('C:\\Users\\anton\\Desktop\\bankstatements\\'),
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    *[Tree(p) for p in (sdl2.dep_bins + glew.dep_bins)],
    [('v', None, 'OPTION')],
    name='BankStatements',
    debug=True,
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
)
