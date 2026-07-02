"""
GUI test script - simulates user interaction
"""
import sys
from pathlib import Path


def test_gui_imports():
    """Test that all GUI components can be imported"""
    print("Testing GUI imports...")

    try:
        import customtkinter as ctk
        print("✓ CustomTkinter imported")

        from ui.main_window import MainWindow
        print("✓ MainWindow imported")

        from converter import PPTToDocConverter, ConversionError
        print("✓ Converter imported")

        print("\n✓ All imports successful!")
        return True

    except Exception as e:
        print(f"\n✗ Import failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_gui_creation():
    """Test that GUI can be created (without mainloop)"""
    print("\nTesting GUI creation...")

    try:
        import customtkinter as ctk
        from ui.main_window import MainWindow

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        app = ctk.CTk()
        window = MainWindow(app)

        print("✓ GUI window created successfully")
        print(f"  Window title: {app.title()}")
        print(f"  Window size: {app.geometry()}")

        # Don't call mainloop in test
        app.destroy()

        return True

    except Exception as e:
        print(f"\n✗ GUI creation failed: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    success1 = test_gui_imports()
    success2 = test_gui_creation()

    if success1 and success2:
        print("\n" + "="*50)
        print("✓ ALL GUI TESTS PASSED")
        print("="*50)
        print("\nThe application is ready to run!")
        print("To launch the GUI, run: python main.py")
        exit(0)
    else:
        print("\n" + "="*50)
        print("✗ SOME TESTS FAILED")
        print("="*50)
        exit(1)
