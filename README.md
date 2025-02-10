# Bifrost - Image Text Extraction Tool 🔍

Bifrost is a user-friendly desktop application that extracts text from images, screenshots, and documents. Whether you're a developer, QA engineer, or just need to grab text from images, Bifrost makes it simple.

## Features ✨

- **Screenshot Capture**: Take screenshots with hotkey `Ctrl+Shift+S`
- **Image Upload**: Support for PNG, JPG, JPEG, BMP, TIFF, and WebP
- **Text Selection**: Select specific areas of images to extract text
- **Image Viewing**:
  - Zoom in/out
  - Pan across image
  - Reset view

## Installation 🚀

1. Download and install [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki)
2. Download the latest Bifrost release
3. Run `bifrost.exe`
4. On first launch, configure Tesseract path (typically `C:\Program Files\Tesseract-OCR\tesseract.exe`)

## Quick Start Guide 📖

1. **Capture Text from Screen**:
   - Press `Ctrl+Shift+S`.
   - Draw a rectangle around the text.
   - Text appears automatically in the application.

2. **Extract from Image File**:
   - Click "Upload Image".
   - Select your image file.
   - Text appears automatically.

3. **Select Specific Text**:
   - Press 'S' to enter selection mode.
   - Draw rectangle around desired text.
   - Selected text is added to current text.

4. **Save Your Results**:
   - Texts are saved in a folder called "Extracted Text".
   - Texts are saved in seperate .txt documents.

## Use Cases 💡

- **Developers**: Extract text from UI mockups or error messages.
- **QA Teams**: Capture and document error messages.
- **Content Creation**: Extract text from images or screenshots.
- **Document Processing**: Convert image-based documents to editable text.
- **Research**: Extract text from figures, charts, or scanned documents.

## System Requirements 💻

- Windows 10 or later
- [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) installed
- 4GB RAM minimum
- 100MB free disk space

## Support 🤝

Found a bug or need help? [Open an issue](../../issues) on our GitHub repository.

## License 📄

MIT License - feel free to use and modify!
