# GitHub Actions Build Guide

## Accessing Pre-Built Executables

### Step-by-Step Instructions:

1. **Go to GitHub Repository**
   ```
   https://github.com/shazam37/ppt_to_docx_converter
   ```

2. **Click on "Actions" Tab**
   - Located at the top of the repository page
   - You'll see all workflow runs

3. **Find the Latest Successful Build**
   - Look for a green checkmark ✓
   - Click on the workflow run name (usually "Build Executables")

4. **Download Your Platform's Executable**
   Scroll down to the "Artifacts" section at the bottom:
   
   - **Windows Users**: Download `PPT-to-DOC-Converter-Windows`
   - **Linux Users**: Download `PPT-to-DOC-Converter-Linux`
   - **macOS Users**: Download `PPT-to-DOC-Converter-macOS`

5. **Extract and Run**
   - Unzip the downloaded file
   - **Windows**: Double-click `PPT-to-DOC-Converter.exe`
   - **Linux/macOS**: Make executable with `chmod +x PPT-to-DOC-Converter` then run `./PPT-to-DOC-Converter`

## Build Status

Check current build status:
```
https://github.com/shazam37/ppt_to_docx_converter/actions
```

## Workflow Details

The CI/CD pipeline automatically:
- ✅ Builds executables for Windows, Linux, and macOS
- ✅ Runs on every push to main branch
- ✅ Can be manually triggered
- ✅ Artifacts are kept for 30 days

## Manual Builds

To trigger a new build manually:
1. Go to Actions tab
2. Select "Build Executables" workflow
3. Click "Run workflow" button
4. Select branch (main)
5. Click green "Run workflow" button

## Build Time

Typical build times:
- Windows: ~3-5 minutes
- Linux: ~3-5 minutes  
- macOS: ~3-5 minutes
- Total: ~5-10 minutes for all platforms

## Troubleshooting

### No Artifacts Available
- Build may still be running (check status)
- Build may have failed (check logs)
- Wait for green checkmark ✓

### Cannot Download Artifact
- You must be logged into GitHub
- Artifacts expire after 30 days
- Trigger a new build if expired

### Build Failed
- Check workflow logs for errors
- Create an issue on GitHub
- Contact repository maintainer

## Direct Links

- **Repository**: https://github.com/shazam37/ppt_to_docx_converter
- **Actions**: https://github.com/shazam37/ppt_to_docx_converter/actions
- **Issues**: https://github.com/shazam37/ppt_to_docx_converter/issues
