# DataMatrix Generator

A web-based application for generating DataMatrix barcodes from manual input or Excel data.

## Overview

The DataMatrix Generator is a browser-based tool that creates DataMatrix (2D) barcodes according to specific formatting rules. The application offers two methods of input:

1. **Manual Input**: Enter a 37-digit code directly
2. **Excel Upload**: Upload an Excel file with specific columns for automatic barcode generation

The generated barcodes can be downloaded as PDF files for printing and practical use.

## Features

- **Dual Input Methods**: Manual entry or Excel file upload
- **Intelligent Column Mapping**: Automatically maps Excel columns to required fields
- **Manual Column Mapping**: Allows custom column mapping when automatic mapping fails
- **Data Validation**: Ensures valid 37-digit barcode strings
- **PDF Export**: Download generated barcodes as perfectly sized PDF files
- **Drag & Drop**: Supports drag and drop for Excel file uploads
- **Responsive Design**: Works on desktop and mobile devices

## Installation

1. Download all files to your web server or local directory:
   - `index.html`
   - `datamatrix.js`

2. Open `index.html` in a web browser

3. No server-side technology is required - this is a completely client-side application

## Usage

### Manual Input Mode

1. Enter a 37-digit code in the text area
2. The input will be validated to ensure it meets the required format
3. The DataMatrix barcode will be generated automatically
4. Click "Download PDF" to save the barcode as a PDF file

### Excel Upload Mode

1. Click the "Excel Upload" tab
2. Either:
   - Click the upload area to select an Excel file
   - Drag and drop an Excel file onto the upload area
3. The application will:
   - Automatically map columns from your Excel file
   - Generate a barcode from the first row of data
   - Display a preview of the generated barcode
4. If the automatic mapping fails, click "Manual Column Mapping" to manually select which columns contain the required information
5. Click "Download PDF" to save the barcode as a PDF file

## Barcode Format

The application generates a 37-digit barcode string in the following format:

| Position | Length | Description            | Format                                                 |
|----------|--------|------------------------|--------------------------------------------------------|
| 1-13     | 13     | ISBN                   | Directly from data                                     |
| 14-17    | 4      | Endsheet Height        | Fixed as "0000"                                        |
| 18-20    | 3      | Spine Size             | 2 digits + trailing zero (e.g., 25 → "250")            |
| 21-24    | 4      | Book Block Height      | 3 digits + trailing zero (e.g., 221 → "2210")          |
| 25-28    | 4      | Trim Off Foot          | Calculated: (BB Height - Trim Height + 3) with zeros   |
| 29-32    | 4      | Trim Height            | 3 digits + trailing zero (e.g., 210 → "2100")          |
| 33-36    | 4      | Trim Width             | 3 digits + trailing zero (e.g., 148 → "1480")          |
| 37       | 1      | Transfer Station       | Fixed as "1"                                           |

### Calculation Rules:

- **Spine Size**:
  - For 2-digit values: Add trailing zero (e.g., 25 → "250")
  - For 1-digit values: Add leading and trailing zero (e.g., 7 → "070")

- **Trim Off Foot**:
  - Calculate as (Book Block Height - Trim Height + 3)
  - For 1-digit result: Add 2 leading zeros and trailing zero (e.g., 5 → "0050")
  - For 2-digit result: Add leading zero and trailing zero (e.g., 14 → "0140")

## Excel File Requirements

The application looks for the following columns in your Excel file:

- **ISBN**: The ISBN or product identification number (13 digits)
- **Height**: The trim height dimension in mm
- **Width**: The trim width dimension in mm
- **Spine**: The spine thickness in mm
- **Cut-Off**: The book block height before trimming in mm

The application will attempt to automatically identify these columns, even if they have different names. Common variations are supported (e.g., "Trim_Height" instead of "Height").

## Troubleshooting

- **Excel Column Mapping Fails**: Use the "Manual Column Mapping" button to manually select which columns contain the required information.
- **PDF Download Not Working**: Ensure you have a valid 37-digit code generated before attempting to download.
- **Barcode Not Generating**: Check the console for detailed error messages (press F12 in most browsers).
- **Browser Compatibility**: The application is built using modern web standards and should work in all current browsers. Internet Explorer is not supported.

## Dependencies

The application uses the following external libraries:

- **DataMatrix.js**: For generating DataMatrix barcodes
- **SheetJS**: For parsing Excel files (`xlsx.full.min.js`)
- **jsPDF**: For PDF generation (`jspdf.umd.min.js`)
- **Bootstrap**: For UI components and responsive design (`bootstrap.bundle.min.js`)
