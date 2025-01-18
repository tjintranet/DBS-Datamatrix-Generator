# DataMatrix Generator

A web-based tool for generating DataMatrix barcodes with precise physical dimensions and specifications.

## Features

- Generate DataMatrix barcodes with exact physical dimensions (17mm x 6mm)
- Customizable barcode content
- Adjustable colors for barcode and background
- Downloads as PDF with precise sizing
- Rectangular format support
- Simple, user-friendly interface
- Integrated link to content calculator

## Usage

### Generating a Barcode

1. Enter your barcode content in the text area
2. Customize appearance if needed:
   - Toggle rectangular format
   - Adjust face color (default: black)
   - Adjust background color (default: white)
3. Preview updates automatically in real-time
4. Click "Download PDF" to get your barcode

### PDF Output Specifications

- Barcode dimensions: 17mm × 6mm
- PDF format: Landscape
- Margins: 10mm on all sides
- Total PDF size: 37mm × 26mm
- High-resolution output
- Centered positioning

### Content Calculator

A link to the DataMatrix Content Calculator is provided below the generator. This calculator helps you format your content correctly for the barcode.

## Technical Details

The application uses:
- Pure JavaScript for barcode generation
- jsPDF for PDF creation
- Bootstrap 5 for the user interface
- SVG for barcode rendering

## Files

- `index.html` - Main application interface
- `datamatrix.js` - Core DataMatrix generation library

## Dependencies

- Bootstrap 5.3.2
- jsPDF 2.5.1
- datamatrix.js (included)

## Hosting

The application can be hosted on any static web server or service like GitHub Pages. No server-side processing is required.

## Credits

DataMatrix library based on:
https://github.com/datalog/datamatrix-svg
Under MIT license
Copyright (c) 2020 Constantine