"""
Build script to create standalone executables for different platforms
"""
import subprocess
import sys
import platform
from pathlib import Path


def build_executable():
    """Build executable for current platform"""
    system = platform.system()
    print(f"Building executable for {system}...")

    # Check if PyInstaller is installed
    try:
        import PyInstaller
    except ImportError:
        print("PyInstaller not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

    # Common PyInstaller arguments
    args = [
        "pyinstaller",
        "--name=PPT-to-DOC-Converter",
        "--onefile",
        "--clean",
        "--noconfirm",
    ]

    # Platform-specific arguments
    if system == "Windows":
        args.extend([
            "--windowed",  # No console window
            "--icon=NONE",
        ])
    elif system == "Darwin":  # macOS
        args.extend([
            "--windowed",
            "--icon=NONE",
        ])
    else:  # Linux
        args.extend([
            "--windowed",
            "--icon=NONE",
        ])

    # Add hidden imports that PyInstaller might miss
    args.extend([
        "--hidden-import=PIL._tkinter_finder",
        "--hidden-import=customtkinter",
        "--hidden-import=pptx",
        "--hidden-import=docx",
    ])

    # Main script
    args.append("main.py")

    print(f"Running: {' '.join(args)}")

    try:
        subprocess.check_call(args)
        print("\n" + "="*60)
        print("✓ Build successful!")
        print("="*60)

        dist_dir = Path("dist")
        if system == "Windows":
            exe_path = dist_dir / "PPT-to-DOC-Converter.exe"
        elif system == "Darwin":
            exe_path = dist_dir / "PPT-to-DOC-Converter.app"
        else:
            exe_path = dist_dir / "PPT-to-DOC-Converter"

        if exe_path.exists():
            print(f"\nExecutable location: {exe_path.absolute()}")
            if exe_path.is_file():
                size_mb = exe_path.stat().st_size / (1024 * 1024)
                print(f"File size: {size_mb:.2f} MB")
        else:
            print(f"\nExecutable should be in: {dist_dir.absolute()}")

        print("\nYou can now distribute the executable to users without Python installed.")

    except subprocess.CalledProcessError as e:
        print(f"\n✗ Build failed: {e}")
        return False

    return True


if __name__ == "__main__":
    success = build_executable()
    exit(0 if success else 1)
