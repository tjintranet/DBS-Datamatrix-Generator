# DataMatrix Generator

A web-based application for generating DataMatrix barcodes from manual input, Excel data, or XML data, specifically designed for book production workflows.

## Overview

The DataMatrix Generator is a browser-based tool that creates DataMatrix (2D) barcodes according to specific formatting rules for book manufacturing. The application offers three methods of input and features intelligent ISBN handling for both limp and cased book formats.

## Features

- **File Upload as Primary Interface**: Excel and XML file processing is the default workflow
- **Multi-Format Support**: Handles Excel (.xlsx, .xls) and XML (.xml) files seamlessly
- **Intelligent ISBN Detection**: Automatically handles both Limp_ISBN and Cased_ISBN columns
- **Dynamic Transfer Station Logic**: Sets position 37 based on book type (1 for Limp, 2 for Cased)
- **Trim Off Head Toggle**: Switch between 7mm (0070) and 3mm (0030) values
- **Smart Column Mapping**: Automatically maps data fields with manual override option
- **Enhanced PDF Export**: Downloads with ISBN-based filenames and trim setting labels
- **Drag & Drop Support**: Easy file uploads for both Excel and XML formats
- **Responsive Design**: Works on desktop and mobile devices

## Installation

1. Download all files to your web server or local directory:
   - `index.html`
   - `datamatrix.js`

2. Open `index.html` in a web browser

3. No server-side technology is required - this is a completely client-side application

## Usage

### File Upload Mode (Primary)

1. The File Upload tab is active by default when you open the application
2. Upload a data file by either:
   - Clicking the upload area to select a file
   - Dragging and dropping a file onto the upload area
3. **Supported file formats:**
   - **Excel files**: `.xlsx`, `.xls`
   - **XML files**: `.xml`
4. The application will:
   - Automatically detect `Limp_ISBN` and `Cased_ISBN` fields
   - Map other required fields (Height, Width, Spine, Cut-Off)
   - Generate a barcode from the data
   - Display a preview of the mapped data
5. If automatic mapping fails, use "Manual Column Mapping" to specify fields
6. Download the barcode as a PDF with intelligent filename

### Manual Input Mode

1. Switch to the "Manual Input" tab
2. Enter a 37-digit code in the text area
3. The input will be validated in real-time
4. The DataMatrix barcode generates automatically
5. Download as PDF when ready

### Trim Off Head Setting

Toggle switch controls the Trim Off Head value (positions 25-28):

- **ON (Default)**: Uses 7mm setting (0070)
- **OFF**: Uses 3mm setting (0030)

The setting automatically regenerates barcodes when file data is loaded.

## Barcode Format

The application generates a 37-digit barcode string in the following format:

| Position | Length | Description            | Format/Rules                                           |
|----------|--------|------------------------|-------------------------------------------------------|
| 1-13     | 13     | ISBN                   | From Limp_ISBN or Cased_ISBN field                    |
| 14-17    | 4      | Endsheet Height        | Fixed as "0000"                                        |
| 18-20    | 3      | Spine Size             | 2 digits + trailing zero OR leading + trailing zero   |
| 21-24    | 4      | Book Block Height      | 3 digits from Cut-Off + trailing zero                 |
| 25-28    | 4      | Trim Off Head          | **7mm = 0070 OR 3mm = 0030 (toggle controlled)**      |
| 29-32    | 4      | Trim Height            | 3 digits from Height + trailing zero                  |
| 33-36    | 4      | Trim Width             | 3 digits from Width + trailing zero                   |
| 37       | 1      | Transfer Station       | **1 = Limp_ISBN used, 2 = Cased_ISBN used**           |

### Key Formatting Rules:

- **ISBN Selection Priority**: Limp_ISBN checked first, then Cased_ISBN if empty
- **Transfer Station Logic**: 
  - `1` if data comes from Limp_ISBN field
  - `2` if data comes from Cased_ISBN field
- **Spine Formatting**:
  - 2-digit values: Add trailing zero (e.g., 24 → "240")
  - 1-digit values: Add leading and trailing zero (e.g., 8 → "080")
- **Trim Off Head**: Fixed toggle-controlled value (no longer calculated)

## File Format Requirements

### Excel Files (.xlsx, .xls)
The application looks for these columns in your Excel file:

#### Required Columns:
- **Limp_ISBN** OR **Cased_ISBN**: 13-digit ISBN (determines Transfer Station)
- **Trim_Height**: Height dimension in mm
- **Trim_Width**: Width dimension in mm  
- **Spine_Size**: Spine thickness in mm
- **Cut_Off**: Book block height before trimming in mm

#### Column Name Variations:
The application recognizes common naming variations:
- Height: `Trim_Height`, `Product_Height`, `Height_MM`
- Width: `Trim_Width`, `Product_Width`, `Width_MM`
- Spine: `Spine_Size`, `Spine_Width`, `Thickness`
- Cut-Off: `Cut_Off`, `Cutoff`, `Bleed`, `Margin`

### XML Files (.xml)
The application expects XML structure with elements containing the required data:

#### Expected XML Structure:
```xml
<?xml version="1.0" encoding="UTF-8"?>
<csv>
  <Wi_Number>324623</Wi_Number>
  <Limp_ISBN>9781739665333</Limp_ISBN>
  <Cased_ISBN></Cased_ISBN>
  <Trim_Height>198</Trim_Height>
  <Trim_Width>129</Trim_Width>
  <Spine_Size>24</Spine_Size>
  <Cut_Off>209</Cut_Off>
  <!-- Additional fields as needed -->
</csv>
```

#### Required XML Elements:
- **Limp_ISBN** OR **Cased_ISBN**: 13-digit ISBN
- **Trim_Height**: Height dimension in mm
- **Trim_Width**: Width dimension in mm
- **Spine_Size**: Spine thickness in mm
- **Cut_Off**: Book block height before trimming in mm

## PDF Export Features

### Enhanced Filename Format:
- **Pattern**: `[13-digit-ISBN]*[trim-setting].pdf`
- **Examples**: 
  - `9781739665333*7mm.pdf` (0070 setting)
  - `9781234567890*3mm.pdf` (0030 setting)

### PDF Contents:
- **DataMatrix barcode** optimally sized for printing
- **Trim setting label** below barcode (8pt font)
- **Professional layout** with proper margins

## Debug Features

- **Debug Panel**: Press `Ctrl+Shift+D` to toggle advanced debugging
- **Column/Field Mapping Display**: View current data field mappings
- **Test Generation**: Test barcode generation with sample data
- **Console Logging**: Detailed process logging in browser console

## Troubleshooting

### File Upload Issues:
- **Column/Field Mapping Fails**: Use "Manual Column Mapping" button
- **No ISBN Found**: Ensure either Limp_ISBN or Cased_ISBN fields contain data
- **Preview Shows Wrong Field**: Check that the correct ISBN field has data
- **XML Parse Error**: Verify XML is well-formed and follows expected structure
- **Unsupported Format**: Ensure file is .xlsx, .xls, or .xml

### Barcode Generation:
- **Invalid Length**: Ensure all required fields are mapped correctly
- **Transfer Station Wrong**: Verify which ISBN field contains data
- **Toggle Not Working**: Only affects file-generated codes, not manual input

### PDF Download:
- **Download Fails**: Ensure a valid 37-digit code is generated first
- **Wrong Filename**: Check that ISBN is properly extracted from barcode
- **Missing Label**: Verify trim setting toggle is working

## Recent Updates

### Version 4.0
- **XML File Support**: Added XML file import alongside Excel support
- **Unified File Interface**: Single upload interface for both Excel and XML files
- **Enhanced File Detection**: Automatic format detection and appropriate parsing
- **Improved Error Handling**: Better feedback for file format issues
- **Updated UI Labels**: Changed from "Excel Upload" to "File Upload"

### Previous Updates (Version 3.0)
- **File Upload Primary**: Made file upload the default workflow
- **Enhanced ISBN Logic**: Added Limp_ISBN/Cased_ISBN detection with Transfer Station logic
- **Improved PDF Export**: Added trim setting labels and ISBN-based filenames
- **Better Preview**: Dynamic field header display based on actual data
- **Enhanced UX**: Improved error handling and user feedback

## Browser Requirements

- Modern web browser with JavaScript enabled
- HTML5 File API support (for file uploads)
- Canvas API support (for PDF generation)
- XML DOM Parser support (for XML file processing)
- No server-side requirements - runs entirely in the browser

## Dependencies

The application uses these external libraries (loaded from CDN):

- **DataMatrix.js**: Custom implementation for DataMatrix barcode generation
- **SheetJS (XLSX)**: Excel file parsing (`xlsx.full.min.js`)
- **jsPDF**: PDF generation (`jspdf.umd.min.js`)
- **Bootstrap 5**: UI components and responsive design (`bootstrap.bundle.min.js`)

## Technical Notes

- **Client-side Processing**: All file parsing and barcode generation happens in the browser
- **Multi-format Support**: Handles both Excel and XML files with unified processing logic
- **No Data Storage**: No user data is stored or transmitted to external servers
- **Cross-platform**: Works on Windows, macOS, Linux, iOS, and Android browsers
- **Performance**: Handles files with complex data structures efficiently
- **XML Processing**: Uses native browser DOM parser for reliable XML handling