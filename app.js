const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

// Function to read JSON files from a directory
function readJSONFilesFromDirectory(directory) {
    const files = fs.readdirSync(directory);
    const jsonFiles = files.filter(file => path.extname(file).toLowerCase() === '.json');
    return jsonFiles.map(file => path.join(directory, file));
}

// Function to create Excel workbook from JSON files
function createExcelFromJSONFiles(jsonFiles) {
    const workbook = new ExcelJS.Workbook();

    jsonFiles.forEach(jsonFile => {
        const jsonData = require(jsonFile);
        const sheetName = path.basename(jsonFile, '.json');

        const sheet = workbook.addWorksheet(sheetName);

        // Add headers
        const headers = Object.keys(jsonData[0]);
        sheet.addRow(headers);

        // Add data
        jsonData.forEach(row => {
            sheet.addRow(Object.values(row));
        });
    });

    return workbook;
}

// Directory containing JSON files
const directory = path.join(__dirname, 'json');

// Read JSON files from the directory
const jsonFiles = readJSONFilesFromDirectory(directory);

// Create Excel workbook from JSON files
const workbook = createExcelFromJSONFiles(jsonFiles);

// Save the workbook to a file
workbook.xlsx.writeFile('output.xlsx')
    .then(() => {
        console.log('Excel file with multiple tabs has been created successfully!');
    })
    .catch(err => {
        console.error('Error:', err);
    });
