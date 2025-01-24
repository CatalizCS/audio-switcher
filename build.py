import PyInstaller.__main__
import os
import shutil


def clean_build():
    """Clean previous build files"""
    dirs_to_clean = ["build", "dist"]
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)


def build_app():
    """Build the application"""
    PyInstaller.__main__.run(
        [
            "audio_switcher.py",
            "--name=AudioSwitcher",
            "--onedir",
            "--windowed",
            "--icon=icon.ico",
            "--add-data=icon.png;.",
            "--add-data=svcl.exe;.",
            "--hidden-import=pystray._win32",
            "--hidden-import=PIL._tkinter_finder",
            "--clean",
            "--noconfirm",
        ]
    )

    # Copy additional files to dist directory
    dist_dir = os.path.join("dist", "AudioSwitcher")
    if not os.path.exists(dist_dir):
        os.makedirs(dist_dir)

    # Create empty config file if it doesn't exist
    config_path = os.path.join(dist_dir, "config.json")
    if not os.path.exists(config_path):
        with open(config_path, "w") as f:
            f.write(
                '{"speakers":[],"headphones":[],"hotkeys":{"switch_device":"ctrl+alt+s","switch_type":"ctrl+alt+t"},"current_type":"Speakers", "kernel_mode_enabled": true, "force_start": false}'
            )


if __name__ == "__main__":
    clean_build()
    build_app()
    print("Build completed successfully!")
