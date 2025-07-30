# CertWizard - Professional Certificate Generator

## Overview
CertWizard is a modern, user-friendly desktop application for generating professional certificates in bulk. It allows users to design certificates using customizable templates and populate them with data from Excel files. The application is built with Python and Tkinter, and supports advanced features such as font selection, color management (RGB/CMYK), drag-and-drop field positioning, and project saving/loading.

## Features
- **Modern GUI**: Built with Tkinter, featuring a clean, modern interface and navigation bar.
- **Template Support**: Load PNG certificate templates (landscape or portrait).
- **Excel Integration**: Import recipient data from `.xlsx` files. Field names are automatically detected from the header row.
- **Customizable Fields**: Each field (e.g., Name, ID, Date) can be toggled, styled (font, size, color), and positioned visually on the template.
- **Font Management**: Supports custom fonts from the `fonts/` directory. Users can add `.ttf` or `.otf` files.
- **Color Spaces**: Choose between RGB and CMYK color spaces for text rendering.
- **Drag-and-Drop**: Place and move text fields directly on the certificate preview.
- **Live Preview**: Instantly preview certificates with sample data before generating.
- **Batch Generation**: Generate certificates for all records in the Excel file, outputting PDFs in organized folders (RGB/CMYK).
- **Project Save/Load**: Save the entire project state (template, field positions, styles, Excel path, etc.) to a `.certwiz` file for later reuse.
- **Progress & Status**: Real-time progress bar and status updates during generation.

## Project Structure
```
certificate-generator/
├── certgen.py                # Main application script
├── requirements.txt          # Python dependencies
├── README.md                 # Project documentation (this file)
├── certgen.ico, certgen.png  # Application icons
├── template(Landscape).png   # Sample certificate template (landscape)
├── template(Potrait).png     # Sample certificate template (portrait)
├── dummy_data.xlsx           # Sample Excel data
├── fonts/                    # Directory for custom font files (.ttf, .otf)
│   └── ...                   # (Add your font files here)
└── ...                      # Other assets and files
```

## Getting Started
### Prerequisites
- Python 3.7+
- pip (Python package manager)

### Installation
1. **Clone or Download the Repository**
2. **Install Dependencies**
   ```sh
   pip install -r requirements.txt
   ```
3. **Add Fonts**
   - Place your desired `.ttf` or `.otf` font files in the `fonts/` directory. The app will auto-detect them.

### Running the Application
```sh
python certgen.py
```

## Usage Guide
1. **Load Template**: Click 'Load Template' and select a PNG certificate template.
2. **Load Excel**: Click 'Load Excel' and select your `.xlsx` file. The first row should contain field names (e.g., Name, ID).
3. **Customize Fields**:
   - Toggle visibility, set font, size, and color for each field.
   - Drag fields on the template to position them.
4. **Preview**: Click 'Preview' to see a sample certificate.
5. **Generate**: Click 'Generate' to create certificates for all records. Choose RGB or CMYK color space and output folder.
6. **Save/Load Project**: Use the 'Project' menu to save or load your project state.

## File Formats
- **Templates**: PNG images (transparent or colored backgrounds supported).
- **Data**: Excel `.xlsx` files. The first row must contain field names.
- **Projects**: `.certwiz` files (JSON format, stores all settings, field positions, and paths).
- **Output**: PDF certificates, named using the first two fields (e.g., `Name_ID_certificate.pdf`).

## Customization & Extensibility
- **Fonts**: Add any number of fonts to the `fonts/` directory. The app will list them for selection.
- **Templates**: Use your own PNG templates. Adjust field positions as needed.
- **Fields**: The app adapts to any field names present in the Excel file.
- **Color Spaces**: Easily switch between RGB and CMYK for professional printing needs.

## Developer Notes
- **Main Entry Point**: `certgen.py` contains the `CertificateApp` class and all logic.
- **UI Structure**: The UI is modular, with separate frames for navigation, settings, status, and canvas.
- **Threading**: Certificate generation runs in a background thread to keep the UI responsive.
- **Image Handling**: Uses Pillow (PIL) for image manipulation and FPDF for PDF output.
- **Platform Support**: Designed for Windows, but should work on Linux/Mac with minor adjustments (icon handling, fonts).
- **Error Handling**: User-friendly error messages and warnings are provided throughout.

## Adding New Features
- To add new field types or data sources, extend the logic in `load_excel` and UI update methods.
- For new export formats, add logic to the `generate_certificates` method.
- To support more image formats, adjust the file dialog filters and image loading code.

## Troubleshooting
- **Fonts Not Detected**: Ensure `.ttf` or `.otf` files are in the `fonts/` directory.
- **Excel Not Loading**: The first row must have field names. Data should start from the second row.
- **Template Not Displaying**: Only PNG files are supported for templates.
- **Output Issues**: Check write permissions for the output directory.

## License
See `LICENSE` for details.

## Credits
- Developed by Shahil SK and contributors.
- Uses [Pillow](https://python-pillow.org/), [openpyxl](https://openpyxl.readthedocs.io/), [FPDF](https://pyfpdf.github.io/), and Tkinter.

---
For questions, suggestions, or contributions, please open an issue or pull request on the repository.
