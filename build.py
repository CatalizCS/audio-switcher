import PyInstaller.__main__
import os
import shutil
import sys

def clean_build():
    """Clean previous build files"""
    dirs_to_clean = ['build', 'dist']
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)

def build_app():
    """Build the application"""
    # Get resource paths
    base_path = os.path.dirname(os.path.abspath(__file__))
    svcl_path = os.path.join(base_path, 'resources', 'svcl.exe')
    icon_path = os.path.join(base_path, 'resources', 'icon.png')

    # Create resources dir if not exists
    os.makedirs(os.path.join(base_path, 'resources'), exist_ok=True)

    # Check resources
    if not os.path.exists(svcl_path):
        print("ERROR: svcl.exe not found in resources folder!")
        print("Please place svcl.exe in the resources folder.")
        sys.exit(1)

    if not os.path.exists(icon_path):
        print("ERROR: icon.png not found in resources folder!")
        print("Please place icon.png in the resources folder.")
        sys.exit(1)

    PyInstaller.__main__.run([
        'audio_switcher.py',
        '--name=AudioSwitcher',
        '--onedir',
        '--windowed',
        '--icon=resources/icon.ico',
        '--add-data=resources/icon.png;resources',
        '--add-data=resources/svcl.exe;resources',
        '--hidden-import=pystray._win32',
        '--hidden-import=PIL._tkinter_finder',
        '--clean',
        '--noconfirm',
        '--uac-admin',
        '--noconsole',
        
    ])

    # Create empty config file
    dist_dir = os.path.join('dist', 'AudioSwitcher')
    config_path = os.path.join(dist_dir, 'config.json')
    if not os.path.exists(config_path):
        with open(config_path, 'w') as f:
            f.write('{"speakers":[],"headphones":[],"hotkeys":{"switch_device":"ctrl+alt+s","switch_type":"ctrl+alt+t"}, "kernel_mode_enabled": true, "force_start": false,"debug_mode":false}')

if __name__ == '__main__':
    clean_build()
    build_app()
    print("Build completed successfully!")
