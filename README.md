# PPT to DOC Converter

A cross-platform desktop application to convert PowerPoint presentations (.pptx) to Word documents (.docx) with **maximum fidelity** and accurate formatting preservation. Designed to create a ditto copy of slides in document format.

## Features

### Core Conversion (Maximum Fidelity)
- **Text Extraction**: Preserves ALL formatting (bold, italic, underline, colors, font sizes, alignment, spacing)
- **Image Embedding**: Extracts and embeds images with original sizing and positioning
- **Table Support**: Converts PowerPoint tables with cell formatting, colors, and structure
- **Layout Preservation**: Multi-column layouts, grouped shapes, natural reading order
- **Bullet Points**: Preserves list formatting, indentation, and nested structures
- **Speaker Notes**: Includes speaker notes from slides
- **Shape Metadata**: Documents shape positions, sizes, and layout information
- **Chart Data**: Attempts to extract chart data as tables

### User Experience
- **Progress Tracking**: Real-time conversion progress with detailed logs
- **Modern GUI**: Dark-themed interface with intuitive controls
- **Batch Ready**: Can be extended for batch conversion
- **Cross-Platform**: Works on Windows, macOS, and Linux
- **Error Handling**: Graceful degradation with detailed error reporting

### Developer Features
- **Programmatic API**: Use as a Python library
- **Callback Support**: Hook into progress and logging
- **Comprehensive Tests**: Full test suite included
- **CI/CD Ready**: GitHub Actions for automated builds

## Installation

### Prerequisites

- Python 3.8 or higher
- pip (Python package manager)

### Setup

1. Clone or download this repository

2. Create a virtual environment:
```bash
python -m venv venv
```

3. Activate the virtual environment:
   - **Windows**: `venv\Scripts\activate`
   - **macOS/Linux**: `source venv/bin/activate`

4. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### GUI Application

Launch the graphical interface:
```bash
python main.py
```

Steps:
1. Click "Browse Input File" to select your PowerPoint file (.pptx)
2. (Optional) Click "Browse Output Location" to choose where to save the Word document
3. Click "Convert to Word" to start the conversion
4. Wait for the conversion to complete - progress will be shown
5. The converted Word document will be saved with all content preserved

### Programmatic Usage

You can also use the converter in your Python code:

```python
from converter import PPTToDocConverter

def progress_callback(percent, message):
    print(f"Progress: {percent}% - {message}")

def log_callback(message):
    print(f"Log: {message}")

converter = PPTToDocConverter(
    progress_cb=progress_callback,
    log_cb=log_callback
)

converter.convert("input.pptx", "output.docx")
```

## Testing

Run the test suite to verify everything works:

```bash
# Basic conversion test
python test_converter.py

# Advanced test with images
python test_advanced.py

# GUI component test
python test_gui.py
```

## What Gets Converted

### Supported Elements

✅ **Text Content**
- Headings and body text
- Text formatting (bold, italic, underline)
- Font colors and sizes
- Paragraph alignment
- Bullet points and numbered lists
- Nested lists

✅ **Images**
- Pictures and photos
- Image sizing (automatically scaled to fit page)
- Image positioning preserved

✅ **Tables**
- Table structure (rows and columns)
- Cell content
- Header row formatting

✅ **Slide Structure**
- Each slide becomes a section in the Word document
- Slide numbers as headings
- Layout names included

### Limitations

❌ **Not Supported**
- Charts (placeholder text added with chart title)
- Animations and transitions
- SmartArt graphics (extracted as text if possible)
- Video/Audio files
- Embedded objects
- Background images
- Slide masters and themes
- Speaker notes (can be added if needed)

## Project Structure

```
ppt-to-doc-converter/
├── converter/
│   ├── __init__.py
│   └── core.py              # Main conversion logic
├── ui/
│   ├── __init__.py
│   └── main_window.py       # GUI interface
├── test_files/              # Generated test files
├── main.py                  # Application entry point
├── test_converter.py        # Basic tests
├── test_advanced.py         # Advanced tests with images
├── test_gui.py              # GUI tests
├── requirements.txt         # Python dependencies
└── README.md               # This file
```

## Dependencies

- **python-pptx**: PowerPoint file reading
- **python-docx**: Word document creation
- **customtkinter**: Modern GUI framework
- **Pillow**: Image processing

## Building Executables

To create standalone executables for distribution:

### Windows
```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "PPT-to-DOC-Converter" main.py
```

### macOS
```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "PPT-to-DOC-Converter" main.py
```

### Linux
```bash
pip install pyinstaller
pyinstaller --onefile --name "PPT-to-DOC-Converter" main.py
```

The executable will be in the `dist/` folder.

## Troubleshooting

### "Cannot open presentation" error
- Ensure the file is a valid .pptx file (not .ppt or corrupted)
- Try opening the file in PowerPoint first to verify it's valid

### Images not appearing
- Check that the original presentation has valid image formats
- Some very large images may need resizing

### Formatting issues
- Complex PowerPoint formatting may not translate perfectly
- Review the output and adjust manually if needed

### GUI doesn't start
- Ensure all dependencies are installed: `pip install -r requirements.txt`
- Check Python version: `python --version` (should be 3.8+)

## Performance

- **Small presentations** (1-10 slides): < 5 seconds
- **Medium presentations** (11-50 slides): 5-30 seconds
- **Large presentations** (50+ slides): 30+ seconds

Processing time depends on:
- Number of slides
- Number and size of images
- Complexity of content (tables, text formatting, etc.)

## License

This project is open source and available for personal and commercial use.

## Contributing

Contributions are welcome! Feel free to submit issues or pull requests.

## Support

For issues, questions, or feature requests, please create an issue in the repository.

## Downloads

### Pre-built Executables

Download the latest executables from [GitHub Actions](https://github.com/shazam37/ppt_to_docx_converter/actions):

1. Go to the [Actions tab](https://github.com/shazam37/ppt_to_docx_converter/actions)
2. Click on the latest successful workflow run
3. Download the executable for your platform:
   - **Windows**: `PPT-to-DOC-Converter-Windows.zip`
   - **Linux**: `PPT-to-DOC-Converter-Linux.zip`
   - **macOS**: `PPT-to-DOC-Converter-macOS.zip`

No Python installation required for pre-built executables!

## Changelog

### Version 2.0.0 (Enhanced Fidelity)
- **Maximum fidelity conversion** - ditto copy of slides
- Multi-column layout preservation
- Shape positioning and metadata
- Speaker notes extraction
- Enhanced text formatting (alignment, spacing, line spacing)
- Chart data extraction
- Natural reading order (sorted by position)
- Comprehensive shape type handling
- Improved error reporting with context

### Version 1.0.0
- Initial release
- Text extraction with formatting
- Image embedding
- Table conversion
- Cross-platform GUI
- Progress tracking
- Comprehensive test suite
