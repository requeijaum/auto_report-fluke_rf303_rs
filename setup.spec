# -*- mode: python -*-

block_cipher = None


a = Analysis(['setup.py', 'r12.py'],
             pathex=['C:\\Users\\HMG\\Dropbox\\rafaelfrequiao\\python_serial_bisturi\\binario_windows\\pyinstaller-funciona'],
             binaries=[],
             datas=[('checklist_bisturi.xlsx', '.')],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='setup',
          debug=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=True )
