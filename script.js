'use strict';

document.addEventListener('DOMContentLoaded', function() {
    // DOM elements
    const box = document.getElementById('box');
    const msg = document.getElementById('msg');
    const dimension = document.getElementById('dimension');
    const padding = document.getElementById('padding');
    const isRectangle = document.getElementById('isRectangle');
    const isOptimized = document.getElementById('isOptimized');
    const faceColor = document.getElementById('faceColor');
    const backgroundColor = document.getElementById('backgroundColor');
    const downloadBtn = document.getElementById('downloadBtn');
    const downloadTxtBtn = document.getElementById('downloadTxtBtn');
    const downloadBatchPDFBtn = document.getElementById('downloadBatchPDFBtn');
    const downloadBatchTXTBtn = document.getElementById('downloadBatchTXTBtn');
    const batchDownloadSection = document.getElementById('batchDownloadSection');
    const batchRowCount = document.getElementById('batchRowCount');
    const batchProgress = document.getElementById('batchProgress');
    const resetBtn = document.getElementById('resetBtn');
    const digitCounter = document.getElementById('digitCounter');
    const fileUploadCard = document.getElementById('fileUploadCard');
    const excelFile = document.getElementById('excelFile');
    const excelDataPreview = document.getElementById('excelDataPreview');
    const excelTableBody = document.getElementById('excelTableBody');
    const generatedCode = document.getElementById('generatedCode');
    const manualMapBtn = document.getElementById('manualMapBtn');
    const mappingModal = new bootstrap.Modal(document.getElementById('mappingModal'));
    const debugPanel = document.getElementById('debugPanel');
    const trimOffToggle = document.getElementById('trimOffToggle');
    const currentTrimOffValue = document.getElementById('currentTrimOffValue');
    
    // Excel data and column mapping
    let excelData = null;
    let availableHeaders = [];
    let columnMapping = {
        isbn: null,
        limpIsbn: null,
        casedIsbn: null,
        height: null,
        width: null,
        spine: null,
        cutOff: null
    };
    
    // Function to calculate Cut_Off based on Trim_Height
    function calculateCutOff(trimHeight) {
        const height = parseFloat(trimHeight);
        
        if (isNaN(height) || height <= 0) {
            console.error("Invalid trim height for Cut_Off calculation:", trimHeight);
            return null;
        }
        
        if (height <= 189) return "200";
        else if (height >= 190 && height <= 198) return "209";
        else if (height > 198 && height <= 210) return "221";
        else if (height > 210 && height <= 216) return "227";
        else if (height > 216 && height <= 229) return "240";
        else if (height > 229 && height <= 235) return "246";
        else if (height > 235 && height <= 244) return "255";
        else if (height > 244 && height <= 246) return "257";
        else if (height > 246 && height <= 254) return "265";
        else if (height > 254 && height <= 276) return "287";
        else if (height > 276 && height <= 279) return "290";
        else if (height > 279 && height <= 297) return "308";
        else {
            console.warn("Trim height outside calculation range:", height);
            return null;
        }
    }
    
    // Function to generate text file content from barcode string
function generateTextFileContent(barcodeString) {
    if (!barcodeString || barcodeString.length !== 37) {
        console.error("Invalid barcode string for text file generation");
        return null;
    }
    
    try {
        const isbn = barcodeString.substring(0, 13);
        const endsheetHeight = barcodeString.substring(13, 17);
        const bookBlockThickness = barcodeString.substring(17, 20);
        const bookBlockHeight = barcodeString.substring(20, 24);
        const trimOffHead = barcodeString.substring(24, 28);
        const finalHeight = barcodeString.substring(28, 32);
        const finalWidth = barcodeString.substring(32, 36);
        const transferStation = barcodeString.substring(36, 37);
        
        console.log("Text file generation from barcode:");
        console.log("ISBN:", isbn);
        console.log("Book_block_thickness:", bookBlockThickness);
        console.log("Book_block_height:", bookBlockHeight);
        console.log("Final_Height:", finalHeight);
        console.log("Final_Width:", finalWidth);
        console.log("Transfer_station_processing:", transferStation);
        
        // For Book_block_height: pad to 5 digits with leading zeros
        const formattedBlockHeightWithZeros = bookBlockHeight.padStart(5, '0');
        
        // For MM_Book_block_height: remove leading zeros, then use the number as-is
        const formattedBlockHeightNoZeros = parseInt(bookBlockHeight, 10).toString();
        
        const textContent = `[job]
Copies=
[files]
File=${isbn}_c.pdf
PFE_Enrich_Sheet=yes
PFE_Set_Page=/u/prismapro/data/shared/PFE3-Covers.properties
[finishing]
Book_block_height=${formattedBlockHeightWithZeros}
Book_block_thickness=${bookBlockThickness}
Endsheet_height=${endsheetHeight}
Final_Height=${finalHeight}
Final_Width=${finalWidth}
Job-ID=${isbn}
MM_Book_block_height=${formattedBlockHeightNoZeros}
Transfer_station_processing=${transferStation}
Job_name=${isbn}_c`;
        
        return textContent;
        
    } catch (error) {
        console.error("Error generating text file content:", error);
        return null;
    }
}
    
    // Function to generate PDF for a single barcode (returns blob)
    async function generatePDFBlob(barcodeString) {
        return new Promise((resolve, reject) => {
            try {
                const tempDiv = document.createElement('div');
                const data = {
                    msg: barcodeString,
                    dim: 256,
                    rct: 1,
                    pad: 2,
                    pal: ['#000000', '#ffffff'],
                    vrb: true
                };
                
                const dataMatrix = DATAMatrix(data);
                tempDiv.appendChild(dataMatrix);
                
                const svg = tempDiv.querySelector('svg');
                if (!svg) {
                    reject(new Error('Failed to generate SVG'));
                    return;
                }
                
                const svgData = new XMLSerializer().serializeToString(svg);
                const canvas = document.createElement('canvas');
                const ctx = canvas.getContext('2d');
                const img = new Image();
                
                img.onload = function() {
                    const { jsPDF } = window.jspdf;
                    const widthScaleFactor = 1.10;
                    const heightScaleFactor = 1.27;
                    const barcodeWidth = 17 * widthScaleFactor;
                    const barcodeHeight = 6 * heightScaleFactor;
                    const margin = 10;
                    const labelHeight = 5;
                    
                    const pageWidth = barcodeWidth + (margin * 2);
                    const pageHeight = barcodeHeight + labelHeight + (margin * 2);
                    
                    const pdf = new jsPDF({
                        orientation: pageWidth > pageHeight ? 'landscape' : 'portrait',
                        unit: 'mm',
                        format: [Math.max(pageWidth, pageHeight), Math.min(pageWidth, pageHeight)]
                    });
                    
                    canvas.width = img.width;
                    canvas.height = img.height;
                    ctx.drawImage(img, 0, 0);
                    
                    const xOffset = (pdf.internal.pageSize.getWidth() - barcodeWidth) / 2;
                    const yOffset = margin;
                    
                    pdf.addImage(
                        canvas.toDataURL('image/png'),
                        'PNG',
                        xOffset,
                        yOffset,
                        barcodeWidth,
                        barcodeHeight
                    );
                    
                    const trimOffValue = getTrimOffHeadValue();
                    const trimOffLabel = trimOffValue === "0070" ? "Trim Off Head: 7mm" : "Trim Off Head: 3mm";
                    
                    pdf.setFontSize(8);
                    pdf.setTextColor(0, 0, 0);
                    const textWidth = pdf.getTextWidth(trimOffLabel);
                    const textX = (pdf.internal.pageSize.getWidth() - textWidth) / 2;
                    const textY = yOffset + barcodeHeight + 3;
                    pdf.text(trimOffLabel, textX, textY);
                    
                    const pdfBlob = pdf.output('blob');
                    resolve(pdfBlob);
                };
                
                img.onerror = function() {
                    reject(new Error('Failed to load image'));
                };
                
                img.src = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svgData)));
                
            } catch (error) {
                reject(error);
            }
        });
    }
    
    // Function to download batch PDFs as ZIP
    async function downloadBatchPDFs() {
        if (!excelData || excelData.length === 0) {
            alert('No data available for batch processing.');
            return;
        }
        
        try {
            const zip = new JSZip();
            const progressBar = batchProgress.querySelector('.progress-bar');
            batchProgress.style.display = 'block';
            
            const trimSuffix = getTrimOffHeadValue() === "0070" ? "*7mm" : "*3mm";
            
            for (let i = 0; i < excelData.length; i++) {
                const row = excelData[i];
                const barcodeString = generateBarcodeString(row);
                
                if (barcodeString) {
                    const isbn = barcodeString.substring(0, 13);
                    const filename = `${isbn}${trimSuffix}.pdf`;
                    
                    console.log(`Generating PDF ${i + 1}/${excelData.length}: ${filename}`);
                    
                    const pdfBlob = await generatePDFBlob(barcodeString);
                    zip.file(filename, pdfBlob);
                    
                    const progress = ((i + 1) / excelData.length) * 100;
                    progressBar.style.width = progress + '%';
                    progressBar.textContent = `${i + 1}/${excelData.length}`;
                }
            }
            
            console.log('Generating ZIP file...');
            const zipBlob = await zip.generateAsync({ type: 'blob' });
            
            const url = window.URL.createObjectURL(zipBlob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `barcodes_batch_${new Date().getTime()}.zip`;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
            
            setTimeout(() => {
                batchProgress.style.display = 'none';
                progressBar.style.width = '0%';
                progressBar.textContent = '';
            }, 1000);
            
            console.log('Batch PDF download complete');
            
        } catch (error) {
            console.error('Error generating batch PDFs:', error);
            alert('Error generating batch PDFs: ' + error.message);
            batchProgress.style.display = 'none';
        }
    }
    
    // Function to download batch TXT files as ZIP
    async function downloadBatchTXTs() {
        if (!excelData || excelData.length === 0) {
            alert('No data available for batch processing.');
            return;
        }
        
        try {
            const zip = new JSZip();
            const progressBar = batchProgress.querySelector('.progress-bar');
            batchProgress.style.display = 'block';
            
            for (let i = 0; i < excelData.length; i++) {
                const row = excelData[i];
                const barcodeString = generateBarcodeString(row);
                
                if (barcodeString) {
                    const isbn = barcodeString.substring(0, 13);
                    const filename = `${isbn}_c.OCT`;
                    const textContent = generateTextFileContent(barcodeString);
                    
                    if (textContent) {
                        console.log(`Adding TXT ${i + 1}/${excelData.length}: ${filename}`);
                        zip.file(filename, textContent);
                        
                        const progress = ((i + 1) / excelData.length) * 100;
                        progressBar.style.width = progress + '%';
                        progressBar.textContent = `${i + 1}/${excelData.length}`;
                    }
                }
            }
            
            console.log('Generating ZIP file...');
            const zipBlob = await zip.generateAsync({ type: 'blob' });
            
            const url = window.URL.createObjectURL(zipBlob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `oct_files_batch_${new Date().getTime()}.zip`;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
            
            setTimeout(() => {
                batchProgress.style.display = 'none';
                progressBar.style.width = '0%';
                progressBar.textContent = '';
            }, 1000);
            
            console.log('Batch TXT download complete');
            
        } catch (error) {
            console.error('Error generating batch TXT files:', error);
            alert('Error generating batch TXT files: ' + error.message);
            batchProgress.style.display = 'none';
        }
    }
    
    // Function to download text file
    function downloadTextFile() {
        let inputText = '';
        const activeTabId = document.querySelector('.tab-pane.active').id;
        
        if (activeTabId === 'manual') {
            inputText = msg.value;
            if (inputText === 'Enter the datamatrix details' || inputText.trim() === '') {
                alert('Please enter data before downloading.');
                return;
            }
            
            if (!validateInput(inputText)) {
                alert('Please enter exactly 37 digits before downloading.');
                return;
            }
        } else {
            inputText = generatedCode.value;
            if (!inputText || inputText.trim() === '') {
                alert('Please upload an Excel file to generate data first.');
                return;
            }
            
            if (inputText.length !== 37) {
                alert('Please generate a valid 37-digit barcode from the Excel data first.');
                return;
            }
        }
        
        const textContent = generateTextFileContent(inputText);
        
        if (!textContent) {
            alert('Error generating text file content.');
            return;
        }
        
        try {
            const isbn = inputText.substring(0, 13);
            const filename = `${isbn}_c.OCT`;
            
            const blob = new Blob([textContent], { type: 'text/plain' });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
            
            console.log(`Text file downloaded: ${filename}`);
            
        } catch (error) {
            console.error('Text file download error:', error);
            alert('Error downloading text file: ' + error.message);
        }
    }
    
    // Function to get current trim off head value based on toggle
    function getTrimOffHeadValue() {
        return trimOffToggle.checked ? "0070" : "0030";
    }
    
    // Function to update trim off display
    function updateTrimOffDisplay() {
        const value = getTrimOffHeadValue();
        currentTrimOffValue.textContent = `Current: ${value}`;
        const label = trimOffToggle.nextElementSibling;
        label.textContent = trimOffToggle.checked ? "Use 0070 (Use on Covers)" : "Use 0030 (Use on Book Blocks)";
    }
    
    // Validation function - checks if input is exactly 37 digits
    function validateInput(input) {
        const digitsOnly = input.replace(/\D/g, '');
        const digitCount = digitsOnly.length;
        
        digitCounter.textContent = `${digitCount}/37 digits`;
        
        if (digitCount === 37) {
            digitCounter.classList.remove('error');
            digitCounter.classList.add('success');
            msg.classList.remove('is-invalid');
            return true;
        } else {
            digitCounter.classList.remove('success');
            digitCounter.classList.add('error');
            if (input !== 'Enter the datamatrix details' && input !== '') {
                msg.classList.add('is-invalid');
            }
            return false;
        }
    }
    
    // Function to pad a number with leading zeros to a specific length
    function padWithZeros(num, targetLength) {
        return num.toString().padStart(targetLength, '0');
    }
    
    // Function to generate barcode string from Excel data
    function generateBarcodeString(row) {
        try {
            console.log("Generating barcode from row:", row);
            console.log("Using column mapping:", columnMapping);
            
            let isbn = '';
            let transferStation = '1';
            
            const limpIsbnValue = columnMapping.limpIsbn ? row[columnMapping.limpIsbn] : '';
            const casedIsbnValue = columnMapping.casedIsbn ? row[columnMapping.casedIsbn] : '';
            
            console.log("Limp_ISBN value:", limpIsbnValue);
            console.log("Cased_ISBN value:", casedIsbnValue);
            
            if (limpIsbnValue && limpIsbnValue.trim() !== '') {
                isbn = limpIsbnValue;
                transferStation = '1';
                console.log("Using Limp_ISBN, Transfer Station = 1");
            } else if (casedIsbnValue && casedIsbnValue.trim() !== '') {
                isbn = casedIsbnValue;
                transferStation = '2';
                console.log("Using Cased_ISBN, Transfer Station = 2");
            } else if (columnMapping.isbn && row[columnMapping.isbn]) {
                isbn = row[columnMapping.isbn];
                transferStation = '1';
                console.log("Using generic ISBN column, Transfer Station = 1");
            }
            
            let height = row[columnMapping.height];
            let width = row[columnMapping.width];
            let spine = row[columnMapping.spine];
            let cutOff = row[columnMapping.cutOff];
            
            if (!cutOff || cutOff === '' || cutOff === null || cutOff === undefined) {
                console.log("Cut_Off is missing, calculating from Trim_Height...");
                cutOff = calculateCutOff(height);
                if (cutOff) {
                    console.log("Calculated Cut_Off:", cutOff);
                } else {
                    console.error("Could not calculate Cut_Off - height may be out of range");
                }
            }
            
            console.log("Raw values:");
            console.log("Final ISBN:", isbn, typeof isbn);
            console.log("Height:", height, typeof height);
            console.log("Width:", width, typeof width);
            console.log("Spine:", spine, typeof spine);
            console.log("Cut-Off:", cutOff, typeof cutOff, "(calculated if was missing)");
            console.log("Transfer Station:", transferStation);
            
            if (isbn === undefined || isbn === null || isbn === '') {
                console.error("No ISBN found in either Limp_ISBN or Cased_ISBN columns");
                return null;
            }
            
            isbn = String(isbn).replace(/\D/g, '');
            
            if (isbn.length > 13) {
                isbn = isbn.substring(0, 13);
            } else if (isbn.length < 13) {
                isbn = isbn.padStart(13, '0');
            }
            
            function cleanMeasurement(value) {
                if (value === undefined || value === null || value === '') return 0;
                if (typeof value === 'string') {
                    value = value.replace(/[^\d.]/g, '');
                }
                return parseFloat(value) || 0;
            }
            
            height = cleanMeasurement(height);
            width = cleanMeasurement(width);
            spine = cleanMeasurement(spine);
            cutOff = cleanMeasurement(cutOff);
            
            console.log("Cleaned measurements:");
            console.log("Height:", height);
            console.log("Width:", width);
            console.log("Spine:", spine);
            console.log("Cut-Off:", cutOff);
            
            const endsheetHeight = "0000";
            
            let spineSegment;
            const spineRounded = Math.round(spine); // Standard rounding: <=0.5 rounds down, >0.5 rounds up
            if (spineRounded >= 10) {
                spineSegment = spineRounded.toString() + "0";
            } else {
                spineSegment = "0" + spineRounded.toString() + "0";
            }
            
            const bbHeight = Math.round(cutOff);
            const bbHeightSegment = padWithZeros(bbHeight, 3) + "0";
            
            const trimOffSegment = getTrimOffHeadValue();
            
            const trimHeightSegment = padWithZeros(Math.round(height), 3) + "0";
            
            const trimWidthSegment = padWithZeros(Math.round(width), 3) + "0";
            
            console.log("Formatted values:");
            console.log("ISBN:", isbn);
            console.log("Endsheet Height:", endsheetHeight);
            console.log("Spine Size:", spineSegment);
            console.log("Book Block Height:", bbHeightSegment);
            console.log("Trim Off Head (FIXED):", trimOffSegment);
            console.log("Trim Height:", trimHeightSegment);
            console.log("Trim Width:", trimWidthSegment);
            console.log("Transfer Station:", transferStation);
            
            const barcodeString = 
                isbn +
                endsheetHeight +
                spineSegment +
                bbHeightSegment +
                trimOffSegment +
                trimHeightSegment +
                trimWidthSegment +
                transferStation;
            
            console.log("Generated barcode string:", barcodeString);
            console.log("Length:", barcodeString.length);
            
            if (barcodeString.length !== 37) {
                console.error("Invalid barcode length:", barcodeString.length);
                return null;
            }
            
            if (/\D/.test(barcodeString)) {
                console.error("Barcode contains non-digits");
                return null;
            }
            
            return barcodeString;
        } catch (error) {
            console.error("Error generating barcode:", error);
            return null;
        }
    }
    
    // Function to find the best column match
    function findBestColumnMatch(headers, target) {
        if (!headers || headers.length === 0) return null;
        
        const targetLower = target.toLowerCase();
        
        const directMatch = headers.find(h => h.toLowerCase() === targetLower);
        if (directMatch) return directMatch;
        
        const partialMatches = headers.filter(h => {
            const headerLower = h.toLowerCase();
            return headerLower.includes(targetLower) || 
                   targetLower.includes(headerLower);
        });
        
        if (partialMatches.length > 0) {
            return partialMatches.sort((a, b) => a.length - b.length)[0];
        }
        
        if (targetLower === 'isbn') {
            const limpIsbn = headers.find(h => h.toLowerCase().includes('limp_isbn'));
            const casedIsbn = headers.find(h => h.toLowerCase().includes('cased_isbn'));
            
            if (limpIsbn) return limpIsbn;
            if (casedIsbn) return casedIsbn;
            
            const isbnVariations = ["isbn13", "ean", "barcode", "isbn_13", "product_id", "id", "code"];
            for (const variation of isbnVariations) {
                const match = headers.find(h => h.toLowerCase().includes(variation.toLowerCase()));
                if (match) return match;
            }
        }
        
        if (targetLower === 'height') {
            const heightVariations = ["trim_height", "product_height", "height_mm", "h", "book_height", "finished_height"];
            for (const variation of heightVariations) {
                const match = headers.find(h => h.toLowerCase().includes(variation.toLowerCase()));
                if (match) return match;
            }
        }
        
        if (targetLower === 'width') {
            const widthVariations = ["trim_width", "product_width", "width_mm", "w", "book_width", "finished_width"];
            for (const variation of widthVariations) {
                const match = headers.find(h => h.toLowerCase().includes(variation.toLowerCase()));
                if (match) return match;
            }
        }
        
        if (targetLower === 'spine') {
            const spineVariations = ["spine_size", "spine_width", "spine_mm", "thickness", "book_thickness", "s"];
            for (const variation of spineVariations) {
                const match = headers.find(h => h.toLowerCase().includes(variation.toLowerCase()));
                if (match) return match;
            }
        }
        
        if (targetLower === 'cutoff') {
            // Be more strict with Cut-Off matching to avoid matching trim_height
            const cutOffVariations = ["cut_off", "cutoff", "cut-off"];
            for (const variation of cutOffVariations) {
                const match = headers.find(h => h.toLowerCase() === variation || h.toLowerCase().endsWith('.' + variation) || h.toLowerCase().endsWith('_' + variation));
                if (match) return match;
            }
        }
        
        return null;
    }
    
    // Update the preview table with Excel data
function updateExcelPreview() {
    if (!excelData || excelData.length === 0) return;
    
    // Check if Cut_Off will be calculated
    const cutOffLabel = columnMapping.cutOff ? 'Cut Off' : 'Cut Off (Calc)';
    
    // Update table header with friendly names instead of column names
    const thead = document.getElementById('excelTableHead');
    thead.innerHTML = `
        <tr>
            <th>ISBN</th>
            <th>Trim Height</th>
            <th>Trim Width</th>
            <th>Spine</th>
            <th>${cutOffLabel}</th>
        </tr>
    `;
    
    // Display first few rows
    let tableHtml = '';
    const maxRows = Math.min(excelData.length, 5);
    
    for (let i = 0; i < maxRows; i++) {
        const row = excelData[i];
        
        // Get ISBN value from whichever column is mapped
        let displayIsbn = '';
        if (columnMapping.limpIsbn && row[columnMapping.limpIsbn]) {
            displayIsbn = row[columnMapping.limpIsbn];
        } else if (columnMapping.casedIsbn && row[columnMapping.casedIsbn]) {
            displayIsbn = row[columnMapping.casedIsbn];
        } else if (columnMapping.isbn && row[columnMapping.isbn]) {
            displayIsbn = row[columnMapping.isbn];
        }
        
        // Show Cut_Off value or calculated value
        let displayCutOff = '';
        if (columnMapping.cutOff && row[columnMapping.cutOff]) {
            displayCutOff = row[columnMapping.cutOff];
        } else {
            // Calculate and show
            const height = row[columnMapping.height];
            displayCutOff = calculateCutOff(height) || 'N/A';
        }
        
        tableHtml += `
            <tr>
                <td>${displayIsbn || ''}</td>
                <td>${row[columnMapping.height] || ''}</td>
                <td>${row[columnMapping.width] || ''}</td>
                <td>${row[columnMapping.spine] || ''}</td>
                <td>${displayCutOff}</td>
            </tr>
        `;
    }
    
    excelTableBody.innerHTML = tableHtml;
}
    
    // Function to parse XML file and convert to JSON format
    function parseXMLFile(xmlContent) {
        try {
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(xmlContent, "text/xml");
            
            const parserError = xmlDoc.getElementsByTagName("parsererror");
            if (parserError.length > 0) {
                throw new Error("XML parsing error: " + parserError[0].textContent);
            }
            
            const rootElement = xmlDoc.documentElement;
            
            // Recursive function to flatten XML structure with dot notation
            function flattenXML(element, prefix = '') {
                const result = {};
                
                // Process child elements
                for (let i = 0; i < element.children.length; i++) {
                    const child = element.children[i];
                    const tagName = child.tagName;
                    const fullPath = prefix ? `${prefix}.${tagName}` : tagName;
                    
                    // Check if this element has child elements or just text
                    if (child.children.length > 0) {
                        // Has children - recurse deeper
                        const nested = flattenXML(child, fullPath);
                        Object.assign(result, nested);
                    } else {
                        // Leaf node - store the value
                        const textContent = child.textContent.trim();
                        result[fullPath] = textContent || '';
                    }
                }
                
                return result;
            }
            
            const jsonObject = flattenXML(rootElement);
            
            console.log("Parsed XML to JSON (flattened):", jsonObject);
            
            // Return as array with single object (to match Excel format)
            return [jsonObject];
            
        } catch (error) {
            console.error("Error parsing XML:", error);
            throw error;
        }
    }
    
    // Function to process and display data from file (Excel or XML)
    function processDataFile(file) {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const fileName = file.name.toLowerCase();
                let jsonData = [];
                
                console.log("Processing file:", file.name);
                
                if (fileName.endsWith('.xml')) {
                    const xmlContent = e.target.result;
                    jsonData = parseXMLFile(xmlContent);
                    
                } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
                    const data = e.target.result;
                    
                    const workbook = XLSX.read(data, {
                        type: 'array',
                        cellDates: true,
                        cellNF: true,
                        raw: false
                    });
                    
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    
                    jsonData = XLSX.utils.sheet_to_json(worksheet, {
                        raw: false,
                        dateNF: 'yyyy-mm-dd',
                        defval: ''
                    });
                } else {
                    throw new Error("Unsupported file format. Please use Excel (.xlsx, .xls) or XML files.");
                }
                
                if (jsonData.length === 0) {
                    alert('No data found in the file.');
                    return;
                }
                
                console.log("Processed data:", jsonData.slice(0, 2));
                
                excelData = jsonData;
                
                availableHeaders = Object.keys(jsonData[0]);
                console.log("Available headers:", availableHeaders);
                
                columnMapping = {
                    isbn: findBestColumnMatch(availableHeaders, 'isbn'),
                    limpIsbn: availableHeaders.find(h => h.toLowerCase().includes('limp_isbn')),
                    casedIsbn: availableHeaders.find(h => h.toLowerCase().includes('cased_isbn')),
                    height: findBestColumnMatch(availableHeaders, 'height'),
                    width: findBestColumnMatch(availableHeaders, 'width'),
                    spine: findBestColumnMatch(availableHeaders, 'spine'),
                    cutOff: findBestColumnMatch(availableHeaders, 'cutoff')
                };
                
                console.log("Automatic column mapping:", columnMapping);
                
                const requiredMappings = ['height', 'width', 'spine'];
                const missingColumns = requiredMappings.filter(key => columnMapping[key] === null);
                
                if (!columnMapping.limpIsbn && !columnMapping.casedIsbn && !columnMapping.isbn) {
                    missingColumns.push('isbn (neither Limp_ISBN nor Cased_ISBN found)');
                }
                
                if (columnMapping.cutOff) {
                    console.log("Cut_Off column found:", columnMapping.cutOff);
                } else {
                    console.log("Cut_Off column NOT found - will be calculated from Trim_Height");
                }
                
                if (missingColumns.length > 0) {
                    console.warn("Missing required columns:", missingColumns);
                    showMappingModal();
                    return;
                }
                
                updateExcelPreview();
                excelDataPreview.classList.remove('d-none');
                
                if (jsonData.length > 1) {
                    batchDownloadSection.style.display = 'block';
                    batchRowCount.textContent = jsonData.length;
                    console.log(`Batch mode enabled: ${jsonData.length} rows detected`);
                } else {
                    batchDownloadSection.style.display = 'none';
                }
                
                const barcodeString = generateBarcodeString(jsonData[0]);
                if (barcodeString) {
                    generatedCode.value = barcodeString;
                    generatedCode.classList.remove('is-invalid');
                    updateDataMatrix();
                } else {
                    generatedCode.value = '';
                    generatedCode.classList.add('is-invalid');
                    console.error("Failed to generate barcode");
                    alert("Could not generate a valid barcode from the file data. Please check the console for details or try manual column mapping.");
                }
                
                updateDebugPanel();
                
            } catch (error) {
                console.error('Error processing file:', error);
                alert('Error processing file: ' + error.message);
            }
        };
        
        reader.onerror = function() {
            alert('Error reading the file.');
        };
        
        if (file.name.toLowerCase().endsWith('.xml')) {
            reader.readAsText(file, 'UTF-8');
        } else {
            reader.readAsArrayBuffer(file);
        }
    }
    
    // Function to show mapping modal with available columns
    function showMappingModal() {
        if (!availableHeaders || availableHeaders.length === 0) {
            alert('No Excel data available. Please upload an Excel file first.');
            return;
        }
        
        const selects = ['isbnColumn', 'heightColumn', 'widthColumn', 'spineColumn', 'cutOffColumn'];
        selects.forEach(selectId => {
            const select = document.getElementById(selectId);
            select.innerHTML = '<option value="">Select column</option>';
            
            availableHeaders.forEach(header => {
                const option = document.createElement('option');
                option.value = header;
                option.textContent = header;
                select.appendChild(option);
            });
            
            const field = selectId.replace('Column', '');
            let mappedValue = columnMapping[field];
            
            if (field === 'isbn') {
                mappedValue = columnMapping.limpIsbn || columnMapping.casedIsbn || columnMapping.isbn;
            }
            
            if (mappedValue) {
                select.value = mappedValue;
            }
        });
        
        mappingModal.show();
    }
    
    // Apply the manual column mapping
    function applyManualMapping() {
        const selectedIsbnColumn = document.getElementById('isbnColumn').value;
        
        columnMapping = {
            isbn: selectedIsbnColumn,
            limpIsbn: selectedIsbnColumn.toLowerCase().includes('limp_isbn') ? selectedIsbnColumn : null,
            casedIsbn: selectedIsbnColumn.toLowerCase().includes('cased_isbn') ? selectedIsbnColumn : null,
            height: document.getElementById('heightColumn').value,
            width: document.getElementById('widthColumn').value,
            spine: document.getElementById('spineColumn').value,
            cutOff: document.getElementById('cutOffColumn').value
        };
        
        console.log("Manual column mapping:", columnMapping);
        
        const missingFields = [];
        if (!columnMapping.isbn) missingFields.push('isbn');
        if (!columnMapping.height) missingFields.push('height');
        if (!columnMapping.width) missingFields.push('width');
        if (!columnMapping.spine) missingFields.push('spine');
        
        if (missingFields.length > 0) {
            alert(`Please select columns for: ${missingFields.join(', ')}`);
            return;
        }
        
        mappingModal.hide();
        
        updateExcelPreview();
        excelDataPreview.classList.remove('d-none');
        
        const barcodeString = generateBarcodeString(excelData[0]);
        if (barcodeString) {
            generatedCode.value = barcodeString;
            generatedCode.classList.remove('is-invalid');
            updateDataMatrix();
        } else {
            generatedCode.value = '';
            generatedCode.classList.add('is-invalid');
            alert('Could not generate a valid barcode string. Please check your data format.');
        }
        
        updateDebugPanel();
    }
    
    // Reset all fields and states
    function resetFields() {
        msg.value = '';
        msg.placeholder = 'Enter the datamatrix details';
        
        msg.classList.remove('is-invalid');
        digitCounter.textContent = '0/37 digits';
        digitCounter.classList.remove('success', 'error');
        
        excelFile.value = '';
        excelDataPreview.classList.add('d-none');
        excelTableBody.innerHTML = '';
        generatedCode.value = '';
        generatedCode.classList.remove('is-invalid');
        excelData = null;
        availableHeaders = [];
        
        batchDownloadSection.style.display = 'none';
        batchProgress.style.display = 'none';
        
        columnMapping = {
            isbn: null,
            limpIsbn: null,
            casedIsbn: null,
            height: null,
            width: null,
            spine: null,
            cutOff: null
        };
        
        trimOffToggle.checked = true;
        updateTrimOffDisplay();
        
        dimension.value = '256';
        padding.value = '2';
        isRectangle.value = 'true';
        isOptimized.value = 'true';
        faceColor.value = '#000000';
        backgroundColor.value = '#ffffff';

        updateDataMatrix();
    }
    
    // Update DataMatrix function
    function updateDataMatrix() {
        while (box.firstChild) {
            box.removeChild(box.firstChild);
        }
        
        let inputText = '';
        const activeTabId = document.querySelector('.tab-pane.active').id;
        
        if (activeTabId === 'manual') {
            inputText = msg.value;
            if (inputText !== 'Enter the datamatrix details' && 
                inputText !== '' && 
                !validateInput(inputText)) {
                return;
            }
        } else {
            inputText = generatedCode.value;
            if (!inputText || inputText.length !== 37) {
                return;
            }
        }

        const data = {
            msg: inputText,
            dim: parseInt(dimension.value),
            rct: 1,
            pad: parseInt(padding.value),
            pal: [faceColor.value, backgroundColor.value],
            vrb: true
        };

        try {
            const dataMatrix = DATAMatrix(data);
            box.appendChild(dataMatrix);
        } catch (error) {
            console.error('Error generating DataMatrix:', error);
        }
    }

    // Download the DataMatrix as PDF
    function downloadPDF() {
        let inputText = '';
        const activeTabId = document.querySelector('.tab-pane.active').id;
        
        if (activeTabId === 'manual') {
            inputText = msg.value;
            if (inputText === 'Enter the datamatrix details' || inputText.trim() === '') {
                alert('Please enter data before downloading.');
                return;
            }
            
            if (!validateInput(inputText)) {
                alert('Please enter exactly 37 digits before downloading.');
                return;
            }
        } else {
            inputText = generatedCode.value;
            if (!inputText || inputText.trim() === '') {
                alert('Please upload an Excel file to generate data first.');
                return;
            }
            
            if (inputText.length !== 37) {
                alert('Please generate a valid 37-digit barcode from the Excel data first.');
                return;
            }
        }
        
        const svg = box.querySelector('svg');
        if (!svg) {
            alert('Please generate a valid Data Matrix first.');
            return;
        }
        
        try {
            const { jsPDF } = window.jspdf;
            
            if (!jsPDF) {
                throw new Error("jsPDF library not loaded");
            }
            
            const svgData = new XMLSerializer().serializeToString(svg);
            
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            
            const img = new Image();
            img.onload = function() {
                const widthScaleFactor = 1.10;
                const heightScaleFactor = 1.27;
                const barcodeWidth = 17 * widthScaleFactor;
                const barcodeHeight = 6 * heightScaleFactor;
                const margin = 10;
                const labelHeight = 5;
                
                const pageWidth = barcodeWidth + (margin * 2);
                const pageHeight = barcodeHeight + labelHeight + (margin * 2);
                
                const pdf = new jsPDF({
                    orientation: pageWidth > pageHeight ? 'landscape' : 'portrait',
                    unit: 'mm',
                    format: [Math.max(pageWidth, pageHeight), Math.min(pageWidth, pageHeight)]
                });
                
                canvas.width = img.width;
                canvas.height = img.height;
                
                ctx.drawImage(img, 0, 0);
                
                const xOffset = (pdf.internal.pageSize.getWidth() - barcodeWidth) / 2;
                const yOffset = margin;
                
                pdf.addImage(
                    canvas.toDataURL('image/png'),
                    'PNG',
                    xOffset,
                    yOffset,
                    barcodeWidth,
                    barcodeHeight
                );
                
                const trimOffValue = getTrimOffHeadValue();
                const trimOffLabel = trimOffValue === "0070" ? "Trim Off Head: 7mm" : "Trim Off Head: 3mm";
                const trimSuffix = trimOffValue === "0070" ? "*7mm" : "*3mm";
                
                pdf.setFontSize(8);
                pdf.setTextColor(0, 0, 0);
                
                const textWidth = pdf.getTextWidth(trimOffLabel);
                const textX = (pdf.internal.pageSize.getWidth() - textWidth) / 2;
                const textY = yOffset + barcodeHeight + 3;
                
                pdf.text(trimOffLabel, textX, textY);
                
                let filenameISBN = '';
                if (inputText && inputText.length >= 13) {
                    filenameISBN = inputText.substring(0, 13);
                } else {
                    filenameISBN = 'datamatrix';
                }
                
                const filename = `${filenameISBN}${trimSuffix}.pdf`;
                
                pdf.save(filename);
            };
            
            img.onerror = function() {
                alert('Error converting SVG to image');
                console.error('Image load error');
            };
            
            img.src = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svgData)));
            
        } catch (error) {
            console.error('PDF generation error:', error);
            alert('Error generating PDF: ' + error.message + '. Check if the jsPDF library is loaded correctly.');
        }
    }
    
    // Update debug panel with current data
    function updateDebugPanel() {
        if (document.getElementById('debugPanel').style.display === 'block') {
            if (excelData && excelData.length > 0) {
                document.getElementById('debugData').value = JSON.stringify(excelData[0], null, 2);
                document.getElementById('debugMapping').value = JSON.stringify(columnMapping, null, 2);
            }
        }
    }
    
    // Event Listeners
    
    trimOffToggle.addEventListener('change', function() {
        updateTrimOffDisplay();
        if (excelData && excelData.length > 0 && 
            Object.values(columnMapping).some(v => v !== null)) {
            const barcodeString = generateBarcodeString(excelData[0]);
            if (barcodeString) {
                generatedCode.value = barcodeString;
                updateDataMatrix();
            }
        }
    });
    
    msg.addEventListener('input', function() {
        validateInput(this.value);
        updateDataMatrix();
    });
    
    fileUploadCard.addEventListener('click', function() {
        excelFile.click();
    });
    
    excelFile.addEventListener('change', function() {
        if (this.files.length > 0) {
            processDataFile(this.files[0]);
        }
    });
    
    fileUploadCard.addEventListener('dragover', function(e) {
        e.preventDefault();
        e.stopPropagation();
        this.style.backgroundColor = '#e9ecef';
    });
    
    fileUploadCard.addEventListener('dragleave', function(e) {
        e.preventDefault();
        e.stopPropagation();
        this.style.backgroundColor = '#f8f9fa';
    });
    
    fileUploadCard.addEventListener('drop', function(e) {
        e.preventDefault();
        e.stopPropagation();
        this.style.backgroundColor = '#f8f9fa';
        
        if (e.dataTransfer.files.length > 0) {
            const file = e.dataTransfer.files[0];
            const fileName = file.name.toLowerCase();
            
            if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls') || fileName.endsWith('.xml')) {
                processDataFile(file);
            } else {
                alert('Please upload an Excel file (.xlsx, .xls) or XML file (.xml)');
            }
        }
    });
                
    manualMapBtn.addEventListener('click', showMappingModal);
    
    document.getElementById('applyMappingBtn').addEventListener('click', applyManualMapping);
    
    document.getElementById('inputTabs').addEventListener('shown.bs.tab', updateDataMatrix);
    
    document.addEventListener('keydown', function(e) {
        if (e.ctrlKey && e.shiftKey && e.key === 'D') {
            e.preventDefault();
            const debugPanel = document.getElementById('debugPanel');
            debugPanel.style.display = debugPanel.style.display === 'none' || debugPanel.style.display === '' ? 'block' : 'none';
            updateDebugPanel();
        }
    });
    
    document.getElementById('testBarcodeBtn').addEventListener('click', function() {
        const testData = {
            'Test_ISBN': document.getElementById('testISBN').value,
            'Test_Height': document.getElementById('testHeight').value,
            'Test_Width': document.getElementById('testWidth').value,
            'Test_Spine': document.getElementById('testSpine').value,
            'Test_CutOff': document.getElementById('testCutOff').value
        };
        
        const testMapping = {
            isbn: 'Test_ISBN',
            limpIsbn: 'Test_ISBN',
            casedIsbn: null,
            height: 'Test_Height',
            width: 'Test_Width',
            spine: 'Test_Spine',
            cutOff: 'Test_CutOff'
        };
        
        const originalMapping = {...columnMapping};
        
        columnMapping = testMapping;
        
        const testBarcode = generateBarcodeString(testData);
        
        columnMapping = originalMapping;
        
        if (testBarcode) {
            alert('Test success! Generated barcode: ' + testBarcode);
        } else {
            alert('Test failed! Check console for details.');
        }
    });
    
    downloadBtn.addEventListener('click', downloadPDF);
    downloadTxtBtn.addEventListener('click', downloadTextFile);
    downloadBatchPDFBtn.addEventListener('click', downloadBatchPDFs);
    downloadBatchTXTBtn.addEventListener('click', downloadBatchTXTs);
    resetBtn.addEventListener('click', resetFields);
    
    msg.addEventListener('focus', function() {
        if (msg.value === 'Enter the datamatrix details') {
            msg.value = '';
            validateInput('');
        }
    });

    msg.addEventListener('blur', function() {
        if (msg.value.trim() === '') {
            msg.value = 'Enter the datamatrix details';
            digitCounter.textContent = '0/37 digits';
            digitCounter.classList.remove('success', 'error');
            msg.classList.remove('is-invalid');
        }
    });

    updateTrimOffDisplay();
    msg.value = 'Enter the datamatrix details';
    digitCounter.textContent = '0/37 digits';
    updateDataMatrix();
});