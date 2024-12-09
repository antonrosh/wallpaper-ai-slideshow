# -*- mode: python ; coding: utf-8 -*-

import os
import site
import ttkbootstrap
import glob
import time
import shutil
from PyInstaller.utils.hooks import collect_all, collect_submodules

# Add version info at the top
VSVersionInfo = {
    'CompanyName': 'Anton Rosh',
    'FileDescription': 'Wallpaper AI Slideshow',
    'FileVersion': '1.0.0',
    'InternalName': 'wallpaper_ai_slideshow',
    'LegalCopyright': '© 2024 Anton Rosh',
    'OriginalFilename': 'wallpaper_ai_slideshow.exe',
    'ProductName': 'Wallpaper AI Slideshow',
    'ProductVersion': '1.0.0'
}

# Find paths
current_dir = os.path.abspath(os.getcwd())
ttkbootstrap_path = os.path.dirname(ttkbootstrap.__file__)
themes_path = os.path.join(ttkbootstrap_path, 'themes')

# Clean build directories
for path in ['build', 'dist']:
    if os.path.exists(path):
        try:
            shutil.rmtree(path)
            time.sleep(1)
        except Exception as e:
            print(f"Warning: Could not clean {path}: {e}")

# Create build directories
build_dir = os.path.join(current_dir, 'build', 'wallpaper_ai_slideshow')
dist_dir = os.path.join(current_dir, 'dist')
os.makedirs(build_dir, exist_ok=True)
os.makedirs(dist_dir, exist_ok=True)

# Collect all dependencies
all_datas = []
all_binaries = []
all_hiddenimports = []

packages_to_collect = [
    'PIL',
    'ttkbootstrap',
    'pystray',
    'cryptography',
]

for pkg in packages_to_collect:
    datas, binaries, hiddenimports = collect_all(pkg)
    all_datas.extend(datas)
    all_binaries.extend(binaries)
    all_hiddenimports.extend(hiddenimports)

# Update icon handling
icon_paths = [
    os.path.join(current_dir, 'app_icon.ico'),
    os.path.join(os.path.dirname(current_dir), 'app_icon.ico'),
]

icon_path = None
for path in icon_paths:
    if os.path.exists(path):
        icon_path = path
        break

if not icon_path:
    print("Warning: Could not find app_icon.ico in any location")

block_cipher = None

a = Analysis(
    ['wallpaper_ai_slideshow.py'],
    pathex=[current_dir],
    binaries=all_binaries,
    datas=[x for x in [
        (icon_path, '.') if icon_path else None,
        (ttkbootstrap_path, 'ttkbootstrap')
    ] if x is not None],
    hiddenimports=all_hiddenimports + [
        'PIL._tkinter_finder',
        'win32api',
        'win32gui',
        'win32con',
        'win32event',
        'win32process',
        'winerror',
        'psutil',
        'logging',
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Filter None values from datas
a.datas = [x for x in a.datas if x is not None]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='wallpaper_ai_slideshow',
    debug=False,  # Change to False for release
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # Change to False to hide console
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_path if icon_path else None,
    version_info=VSVersionInfo
)

# Add collection with error handling
try:
    coll = COLLECT(
        exe,
        a.binaries,
        a.zipfiles,
        a.datas,
        strip=False,
        upx=True,
        upx_exclude=[],
        name='wallpaper_ai_slideshow'
    )
except Exception as e:
    print(f"Error during collection: {e}")
    raise
