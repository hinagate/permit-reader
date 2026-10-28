# PyInstaller spec for PermitReader.exe - build with build.bat
from PyInstaller.utils.hooks import collect_data_files

a = Analysis(
    ["permit_reader.py"],
    datas=collect_data_files("rapidocr"),   # OCR models and config files
    hiddenimports=["rapidocr.main"],         # loaded lazily by rapidocr/__init__.py
    excludes=["tkinter", "PIL._avif"],
)
# OpenCV's video backend (FFmpeg) is never used and adds ~30 MB.
a.binaries = [b for b in a.binaries if "opencv_videoio_ffmpeg" not in b[0]]

pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="PermitReader",
    console=True,
    upx=False,
)
