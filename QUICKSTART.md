# Quick Start Guide

## For End Users (Non-Technical)

### Option 1: Use the Pre-built Executable (Easiest)

If you received a pre-built executable:

1. **Windows**: Double-click `PPT-to-DOC-Converter.exe`
2. **macOS**: Double-click `PPT-to-DOC-Converter.app`
3. **Linux**: Make executable with `chmod +x PPT-to-DOC-Converter`, then run `./PPT-to-DOC-Converter`

### Option 2: Run from Source

If you have Python installed:

1. Open a terminal/command prompt in this folder
2. Run: `python main.py`
3. The GUI will open automatically

## Using the Application

1. Click **"Browse Input File"** and select your PowerPoint file (.pptx)
2. Click **"Convert to Word"** button
3. Wait for conversion to complete (progress bar shows status)
4. The Word document will be saved in the same folder as your PowerPoint file

## For Developers

### Setup Development Environment

```bash
# Create virtual environment
python -m venv venv

# Activate it
source venv/bin/activate  # On Linux/Mac
# or
venv\Scripts\activate  # On Windows

# Install dependencies
pip install -r requirements.txt
```

### Run Tests

```bash
# Run all tests
python run_all_tests.py

# Or run individual tests
python test_converter.py
python test_advanced.py
python test_gui.py
```

### Build Executable

```bash
python build_executable.py
```

The executable will be in the `dist/` folder.

### Project Structure

- `main.py` - Entry point for the application
- `converter/core.py` - Core conversion logic
- `ui/main_window.py` - GUI interface
- `test_*.py` - Test scripts
- `requirements.txt` - Python dependencies

## Supported Features

✅ Text with formatting (bold, italic, colors, sizes)
✅ Images (embedded in Word document)
✅ Tables (converted to Word tables)
✅ Bullet points and lists
✅ Multiple text boxes per slide

❌ Charts (placeholder added)
❌ Animations
❌ Video/Audio

## Common Issues

**Problem**: "Cannot open presentation"
**Solution**: Make sure your file is .pptx format (not .ppt). Try opening it in PowerPoint first.

**Problem**: GUI doesn't start
**Solution**: Install dependencies with `pip install -r requirements.txt`

**Problem**: Images missing
**Solution**: Check that images in the original presentation are valid

## Performance

- **Small files** (1-10 slides): < 5 seconds
- **Medium files** (11-50 slides): 5-30 seconds
- **Large files** (50+ slides): 30+ seconds

## Need Help?

Check the full [README.md](README.md) for detailed documentation.
