"""
Master test script - runs all tests to verify the application
"""
import subprocess
import sys
from pathlib import Path


def run_test(test_file, description):
    """Run a test script and return success status"""
    print("\n" + "="*70)
    print(f"Running: {description}")
    print("="*70)

    try:
        result = subprocess.run(
            [sys.executable, test_file],
            capture_output=False,
            text=True,
            check=True
        )
        print(f"\n✓ {description} - PASSED")
        return True
    except subprocess.CalledProcessError:
        print(f"\n✗ {description} - FAILED")
        return False


def main():
    """Run all tests"""
    print("PPT to DOC Converter - Comprehensive Test Suite")
    print("="*70)

    tests = [
        ("test_gui.py", "GUI Component Tests"),
        ("test_converter.py", "Basic Conversion Tests"),
        ("test_advanced.py", "Advanced Conversion Tests (with images)"),
    ]

    results = []
    for test_file, description in tests:
        if not Path(test_file).exists():
            print(f"\n✗ Test file not found: {test_file}")
            results.append(False)
            continue

        results.append(run_test(test_file, description))

    # Summary
    print("\n" + "="*70)
    print("TEST SUMMARY")
    print("="*70)

    passed = sum(results)
    total = len(results)

    for i, (test_file, description) in enumerate(tests):
        status = "✓ PASS" if results[i] else "✗ FAIL"
        print(f"{status} - {description}")

    print("="*70)
    print(f"Total: {passed}/{total} tests passed")

    if passed == total:
        print("\n🎉 ALL TESTS PASSED! The application is ready to use.")
        print("\nTo run the application:")
        print("  python main.py")
        print("\nTo build an executable:")
        print("  python build_executable.py")
        return True
    else:
        print(f"\n⚠️  {total - passed} test(s) failed. Please review the errors above.")
        return False


if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)
