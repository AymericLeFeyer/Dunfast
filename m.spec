# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['m.py'],
             pathex=['C:\\Users\\Aymeric\\Documents\\GitHub\\Dunfast'],
             binaries=[],
             datas=[],
             hiddenimports=['pyexcel.plugins.sources.file_input', 'pyexcel.plugins.parsers.excel', 'pyexcel_xls', 'pyexcel_xls.xlsr'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='m',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=True )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='m')
