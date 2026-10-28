# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

import PyQt6

project_root = Path(SPECPATH)
datas = [(str(project_root / "samples"), "samples")]
qt_bin = Path(PyQt6.__file__).resolve().parent / "Qt6" / "bin"
qt_runtime_names = (
    "MSVCP140.dll",
    "MSVCP140_1.dll",
    "MSVCP140_2.dll",
    "VCRUNTIME140.dll",
    "VCRUNTIME140_1.dll",
)

for qt_runtime_name in qt_runtime_names:
    if not (qt_bin / qt_runtime_name).is_file():
        raise FileNotFoundError(f"PyQt6 runtime is missing {qt_runtime_name}: {qt_bin}")

a = Analysis(
    [str(project_root / "tafel_analyzer.py")],
    pathex=[str(project_root)],
    binaries=[],
    datas=datas,
    # Figure.savefig() imports vector backends dynamically by file format, so
    # PyInstaller cannot infer these from the Qt canvas imports alone.
    hiddenimports=[
        "matplotlib.backends.backend_pdf",
        "matplotlib.backends.backend_svg",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

# Python 3.12 contributes MSVC runtime DLLs to the bundle root. Qt resolves
# DLLs already loaded by basename, so retain the matching copies that ship with
# PyQt6 at the root as well as in Qt6/bin. Otherwise an older redistributable
# can be selected before QtWidgets is imported.
qt_runtime_sources = {
    runtime_name.lower(): str(qt_bin / runtime_name)
    for runtime_name in qt_runtime_names
}
replaced_runtime_names = set()
patched_binaries = []
for destination, source, typecode in a.binaries:
    runtime_source = qt_runtime_sources.get(destination.lower())
    if runtime_source is not None:
        patched_binaries.append((destination, runtime_source, typecode))
        replaced_runtime_names.add(destination.lower())
    else:
        patched_binaries.append((destination, source, typecode))

for runtime_name, runtime_source in qt_runtime_sources.items():
    if runtime_name not in replaced_runtime_names:
        patched_binaries.append((runtime_name, runtime_source, "BINARY"))

# PyInstaller can resolve an unrelated Poppler ICU from the build host PATH
# into the bundle root. Qt expects Windows' intrinsic unversioned ICU API, so
# leave this root destination absent and preserve every other collected binary.
a.binaries = [
    binary
    for binary in patched_binaries
    if binary[0].casefold() != "icuuc.dll"
]
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="TafelAnalyzer",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
)
