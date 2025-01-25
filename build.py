import PyInstaller.__main__
import os
import shutil
import sys


def clean_build():
    """Clean previous build files"""
    dirs_to_clean = ["build", "dist", "temp_resources"]
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)


def build_app():
    """Build the application"""
    base_path = os.path.dirname(os.path.abspath(__file__))
    dist_dir = os.path.join("dist", "AudioSwitcher")
    temp_resources = os.path.join("temp_resources")

    # Clean and create temp directory
    if os.path.exists(temp_resources):
        shutil.rmtree(temp_resources)
    os.makedirs(temp_resources)

    # Look for resources
    resource_files = {"svcl.exe": None, "icon.png": None}

    # Search paths for resources
    search_paths = [base_path, os.path.join(base_path, "resources"), os.getcwd()]

    # Find resources
    for resource in resource_files:
        for path in search_paths:
            full_path = os.path.join(path, resource)
            if os.path.exists(full_path):
                resource_files[resource] = full_path
                print(f"Found {resource} at: {full_path}")
                break

        if not resource_files[resource]:
            print(f"ERROR: {resource} not found in any of these locations:")
            for path in search_paths:
                print(f"  - {os.path.join(path, resource)}")
            sys.exit(1)

    # Copy resources to temp directory
    for resource, path in resource_files.items():
        dest = os.path.join(temp_resources, resource)
        shutil.copy2(path, dest)
        print(f"Copied {resource} to temp directory")

    # Build with temp resources
    PyInstaller.__main__.run(
        [
            "audio_switcher.py",
            "--name=AudioSwitcher",
            "--onedir",
            "--windowed",
            "--icon=resources/icon.ico",
            f"--add-data={temp_resources}/*;resources/",
            "--hidden-import=pystray._win32",
            "--hidden-import=PIL._tkinter_finder",
            "--clean",
            "--noconfirm",
            "--uac-admin",
            "--noconsole",
            "--workpath=build",
            "--distpath=dist",
        ]
    )

    # Create final resources directory and copy files
    final_resources = os.path.join(dist_dir, "resources")
    os.makedirs(final_resources, exist_ok=True)

    for resource in resource_files:
        src = os.path.join(temp_resources, resource)
        dst = os.path.join(final_resources, resource)
        if os.path.exists(src):
            shutil.copy2(src, dst)
            print(f"Copied {resource} to final location: {dst}")

    # Create config file
    config_path = os.path.join(dist_dir, "config.json")
    if not os.path.exists(config_path):
        with open(config_path, "w") as f:
            f.write(
                '{"speakers":[],"headphones":[],"hotkeys":{"switch_device":"ctrl+alt+s","switch_type":"ctrl+alt+t"}, "kernel_mode_enabled": true, "force_start": false,"debug_mode":false}'
            )

    # Clean up temp directory
    shutil.rmtree(temp_resources)

    print("\nBuild completed successfully!")
    print(f"Application files are in: {os.path.abspath(dist_dir)}")
    print("\nFiles in distribution:")
    for root, dirs, files in os.walk(dist_dir):
        level = root.replace(dist_dir, "").count(os.sep)
        indent = "  " * level
        print(f"{indent}{os.path.basename(root)}/")
        for f in files:
            print(f"{indent}  {f}")


if __name__ == "__main__":
    clean_build()
    build_app()
