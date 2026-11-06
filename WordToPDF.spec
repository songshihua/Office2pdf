# Build from any checkout location; no machine-specific paths.
from pathlib import Path

root = Path(SPECPATH)
a = Analysis(
    [str(root / 'src' / 'word_pdf_app.py')],
    pathex=[str(root / 'src')],
    binaries=[], datas=[], hiddenimports=[], hookspath=[], hooksconfig={},
    runtime_hooks=[],
    excludes=['numpy', 'scipy', 'matplotlib', 'pandas', 'skimage', 'IPython',
              'tkinter', 'torch', 'sympy', 'PyQt5.QtWebEngineWidgets',
              'PyQt5.QtWebEngineCore', 'PyQt5.QtQml', 'PyQt5.QtQuick'],
    noarchive=False, optimize=0,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [], name='Word转PDF',
    debug=False, bootloader_ignore_signals=False, strip=False, upx=False,
    console=False, disable_windowed_traceback=False,
    icon=[str(root / 'docs' / 'assets' / 'app.ico')],
)
