# -*- mode: python -*-

block_cipher = None


a = Analysis(['r15.py'],
             pathex=['C:\\Users\\HMG\\Dropbox\\rafaelfrequiao\\python_serial_bisturi\\binario_windows\\pyinstaller-funciona'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
			 
			 
a.datas += [("checklist_bisturi.xlsx" , "C:\\Users\\HMG\\Dropbox\\rafaelfrequiao\\python_serial_bisturi\\binario_windows\\pyinstaller-funciona\checklist_bisturi.xlsx" , ".")]
			 
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='Relatorio - FLUKE RS 303 RS',
          debug=False,
          strip=False,
          upx=True,
		  icon=".\\icone.ico",
          runtime_tmpdir=None,
          console=True )
