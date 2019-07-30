# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['excelToGraph.py'],
             pathex=['C:\\Users\\kdfri\\Documents\\GitRepos\\GraphViz_Tutorial'],
             binaries=[],
             datas=[ ('Graphviz2.38', 'Graphviz2.38'),
                      ('images','images'),
                      ('config','config')],
             hiddenimports=[],
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
          name='excelToGraph',
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
               name='excelToGraph')
