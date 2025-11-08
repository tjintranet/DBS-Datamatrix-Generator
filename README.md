# DataMatrix Generator

A web-based application for generating DataMatrix barcodes and OCT text files from manual input, Excel data, or XML data, specifically designed for book production workflows with full batch processing support.

## Overview

The DataMatrix Generator is a browser-based tool that creates DataMatrix (2D) barcodes and OCT configuration files according to specific formatting rules for book manufacturing. The application offers multiple input methods, intelligent ISBN handling for both limp and cased book formats, and powerful batch processing capabilities for high-volume production.

## Features

### Core Functionality
- **File Upload as Primary Interface**: Excel and XML file processing is the default workflow
- **Multi-Format Support**: Handles Excel (.xlsx, .xls) and XML (.xml) files seamlessly
- **Intelligent ISBN Detection**: Automatically handles both Limp_ISBN and Cased_ISBN columns
- **Dynamic Transfer Station Logic**: Sets position 37 based on book type (1 for Limp, 2 for Cased)
- **Trim Off Head Toggle**: Switch between 7mm (0070) and 3mm (0030) values
- **Smart Column Mapping**: Automatically maps data fields with manual override option
- **Auto Cut-Off Calculation**: Automatically calculates Cut-Off values when missing from input data

### Output Options
- **PDF Barcode Export**: Downloads DataMatrix barcodes as PDFs with ISBN-based filenames and trim setting labels
- **OCT Text File Export**: Generates job configuration files in OCT format for printing systems
- **Batch Processing**: Process multiple rows at once with ZIP file downloads
- **Drag & Drop Support**: Easy file uploads for both Excel and XML formats

### Batch Processing
- **Automatic Batch Detection**: Detects multi-row files and enables batch mode
- **Batch PDF Generation**: Download all barcodes as a single ZIP file
- **Batch OCT Generation**: Download all text files as a single ZIP file
- **Progress Tracking**: Visual progress bar shows generation status
- **Smart Filename Generation**: Each file automatically named with ISBN and settings

### User Experience
- **Responsive Design**: Works on desktop and mobile devices
- **Real-time Validation**: Live feedback on data entry
- **Preview Mode**: See mapped data before generation
- **Debug Tools**: Advanced debugging panel (Ctrl+Shift+D)

## Installation

1. Download all files to your web server or local directory:
   - `index.html`
   - `script.js`
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
   - Calculate Cut-Off automatically if missing
   - Generate a barcode from the data
   - Display a preview of the mapped data
5. If automatic mapping fails, use "Manual Column Mapping" to specify fields
6. Choose your download option:
   - **Single Item**: Download PDF or TXT for the first row only
   - **Batch Mode** (if multiple rows): Download all items as ZIP files

### Batch Processing Workflow

1. **Upload Multi-Row Excel/XML File**
   - File must contain multiple rows of data
   - Each row represents one book/barcode

2. **Batch Mode Activates Automatically**
   - Blue alert shows: "Batch Mode: X rows detected"
   - Batch download buttons appear below single download buttons

3. **Generate Batch**
   - Click "Download All PDFs (ZIP)" for all barcodes
   - Click "Download All TXTs (ZIP)" for all OCT files
   - Progress bar shows generation progress
   - ZIP file downloads automatically when complete

4. **Single Item Downloads Still Available**
   - Original "Download PDF" and "Download TXT" buttons work for first row
   - Use these for testing or single-item processing

### Manual Input Mode

1. Switch to the "Manual Input" tab
2. Enter a 37-digit code in the text area
3. The input will be validated in real-time
4. The DataMatrix barcode generates automatically
5. Download as PDF or TXT when ready

### Trim Off Head Setting

Toggle switch controls the Trim Off Head value (positions 25-28):

- **ON (Default)**: Uses 7mm setting (0070) - Use on Covers
- **OFF**: Uses 3mm setting (0030) - Use on Book Blocks

The setting automatically regenerates barcodes when file data is loaded.

## Barcode Format

The application generates a 37-digit barcode string in the following format:

| Position | Length | Description            | Format/Rules                                           |
|----------|--------|------------------------|-------------------------------------------------------|
| 1-13     | 13     | ISBN                   | From Limp_ISBN or Cased_ISBN field                    |
| 14-17    | 4      | Endsheet Height        | Fixed as "0000"                                        |
| 18-20    | 3      | Spine Size             | Standard rounding applied, then formatted with zeros   |
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
- **Spine Formatting & Rounding**:
  - **Rounding**: Standard mathematical rounding (≤0.5 rounds down, >0.5 rounds up)
    - 8.4mm → 8mm
    - 8.5mm → 9mm (rounds up at exactly 0.5)
    - 8.6mm → 9mm
  - **After rounding**: 2-digit values add trailing zero (e.g., 24 → "240")
  - **After rounding**: 1-digit values add leading and trailing zero (e.g., 8 → "080")
- **Trim Off Head**: Fixed toggle-controlled value (no longer calculated)
- **Cut-Off Auto-Calculation**: If missing, calculated based on Trim Height using predefined ranges

## OCT Text File Format

The application generates OCT configuration files with the following format:

```
[job]
Copies=
[files]
File=[ISBN]_c.pdf
PFE_Enrich_Sheet=yes
PFE_Set_Page=/u/prismapro/data/shared/PFE3-Covers.properties
[finishing]
Book_block_height=[5 digits from positions 21-24]
Book_block_thickness=[3 digits from positions 18-20]
Endsheet_height=[4 digits from positions 14-17, always 0000]
Final_Height=[4 digits from positions 29-32]
Final_Width=[4 digits from positions 33-36]
Job-ID=[13 digits from positions 1-13]
MM_Book_block_height=[Book_block_height - no leading zero]
Transfer_station_processing=[1 digit from position 37]
Job_name=[ISBN]_c
```

### OCT File Naming
- **Filename Pattern**: `[ISBN]_c.OCT`
- **Example**: `9781234567890_c.OCT`

## Cut-Off Auto-Calculation

When the Cut-Off field is missing from uploaded data, it is automatically calculated based on Trim Height:

| Trim Height Range (mm) | Calculated Cut-Off (mm) |
|------------------------|-------------------------|
| ≤ 189                  | 200                     |
| 190 - 198              | 209                     |
| 199 - 210              | 221                     |
| 211 - 216              | 227                     |
| 217 - 229              | 240                     |
| 230 - 235              | 246                     |
| 236 - 244              | 255                     |
| 245 - 246              | 257                     |
| 247 - 254              | 265                     |
| 255 - 276              | 287                     |
| 277 - 279              | 290                     |
| 280 - 297              | 308                     |

## File Format Requirements

### Excel Files (.xlsx, .xls)
The application looks for these columns in your Excel file:

#### Required Columns:
- **Limp_ISBN** OR **Cased_ISBN**: 13-digit ISBN (determines Transfer Station)
- **Trim_Height**: Height dimension in mm
- **Trim_Width**: Width dimension in mm  
- **Spine_Size** (or BULK): Spine thickness in mm (decimals will be rounded using standard rounding)
- **Cut_Off**: Book block height before trimming in mm (OPTIONAL - will be calculated if missing)

#### Column Name Variations:
The application recognizes common naming variations:
- Height: `Trim_Height`, `Product_Height`, `Height_MM`
- Width: `Trim_Width`, `Product_Width`, `Width_MM`
- Spine: `Spine_Size`, `BULK`, `Spine_Width`, `Thickness`
- Cut-Off: `Cut_Off`, `Cutoff`, `Bleed`, `Margin` (optional)

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
- **Cut_Off**: Book block height in mm (optional - will be calculated if missing)

## Download Options

### Single Item Downloads

#### PDF Export
- **Filename Format**: `[ISBN]*[trim-setting].pdf`
- **Examples**: 
  - `9781739665333*7mm.pdf` (0070 setting)
  - `9781234567890*3mm.pdf` (0030 setting)
- **Contents**: DataMatrix barcode with trim setting label

#### TXT/OCT Export
- **Filename Format**: `[ISBN]_c.OCT`
- **Example**: `9781234567890_c.OCT`
- **Contents**: Complete job configuration in OCT format

### Batch Downloads

#### Batch PDF Export (ZIP)
- **Filename Format**: `barcodes_batch_[timestamp].zip`
- **Contains**: Individual PDF files for each row, each named `[ISBN]*[trim].pdf`
- **Progress**: Visual progress bar shows generation status

#### Batch TXT Export (ZIP)
- **Filename Format**: `oct_files_batch_[timestamp].zip`
- **Contains**: Individual OCT files for each row, each named `[ISBN]_c.OCT`
- **Progress**: Visual progress bar shows generation status

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
- **Spine Rounding Issues**: Spine values use standard rounding (≤0.5 down, >0.5 up)

### Batch Processing:
- **Batch Mode Not Showing**: Ensure Excel file has more than one row of data
- **Some Items Failing**: Check console for errors on specific rows
- **ZIP Download Fails**: Ensure all rows have valid data
- **Progress Bar Stuck**: Check browser console for errors

### PDF/TXT Download:
- **Download Fails**: Ensure a valid 37-digit code is generated first
- **Wrong Filename**: Check that ISBN is properly extracted from barcode
- **Missing Label**: Verify trim setting toggle is working

## Recent Updates

### Version 5.0 - Batch Processing
- **Batch PDF Generation**: Download all barcodes as ZIP
- **Batch TXT Generation**: Download all OCT files as ZIP
- **Progress Tracking**: Visual feedback during batch processing
- **Smart Detection**: Automatic batch mode activation for multi-row files
- **JSZip Integration**: Efficient ZIP file creation

### Version 4.0 - OCT File Support
- **TXT File Generation**: Create OCT configuration files
- **Dual Download Options**: Separate buttons for PDF and TXT downloads
- **Text File Formatting**: Proper OCT format with all job parameters

### Version 3.0 - Enhanced File Processing
- **XML File Support**: Added XML file import alongside Excel support
- **Unified File Interface**: Single upload interface for both Excel and XML files
- **Enhanced File Detection**: Automatic format detection and appropriate parsing
- **Cut-Off Auto-Calculation**: Automatic calculation when Cut-Off is missing
- **Improved ISBN Logic**: Added Limp_ISBN/Cased_ISBN detection with Transfer Station logic
- **Better Preview**: Dynamic field header display based on actual data
- **Enhanced UX**: Improved error handling and user feedback

## Browser Requirements

- Modern web browser with JavaScript enabled
- HTML5 File API support (for file uploads)
- Canvas API support (for PDF generation)
- XML DOM Parser support (for XML file processing)
- Blob API support (for file downloads)
- No server-side requirements - runs entirely in the browser

## Dependencies

The application uses these external libraries (loaded from CDN):

- **DataMatrix.js**: Custom implementation for DataMatrix barcode generation
- **JSZip (3.10.1)**: ZIP file creation for batch downloads
- **SheetJS (XLSX 0.18.5)**: Excel file parsing (`xlsx.full.min.js`)
- **jsPDF (2.5.1)**: PDF generation (`jspdf.umd.min.js`)
- **Bootstrap 5 (5.3.2)**: UI components and responsive design (`bootstrap.bundle.min.js`)

## Technical Notes

- **Client-side Processing**: All file parsing and barcode generation happens in the browser
- **Multi-format Support**: Handles both Excel and XML files with unified processing logic
- **No Data Storage**: No user data is stored or transmitted to external servers
- **Cross-platform**: Works on Windows, macOS, Linux, iOS, and Android browsers
- **Performance**: Handles large batches efficiently with async processing
- **Memory Efficient**: Sequential batch processing to avoid memory issues
- **XML Processing**: Uses native browser DOM parser for reliable XML handling

## Example Workflow

### Single Item:
1. Upload Excel with 1 book
2. Preview shows mapped data
3. Click "Download PDF" → Receive `9781234567890*7mm.pdf`
4. Click "Download TXT" → Receive `9781234567890_c.OCT`

### Batch Processing:
1. Upload Excel with 50 books
2. See "Batch Mode: 50 rows detected"
3. Click "Download All PDFs (ZIP)"
4. Watch progress: "1/50", "2/50", ... "50/50"
5. Receive `barcodes_batch_1234567890.zip` with 50 PDF files
6. Click "Download All TXTs (ZIP)"
7. Receive `oct_files_batch_1234567890.zip` with 50 OCT files

## Support

For issues or questions, check the browser console (F12) for detailed error messages and logging information.