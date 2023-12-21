const fs = require('fs');
const XLSX = require('xlsx');

function processJSON(jsonData, sheets, prefix = '') {
  if (!jsonData || typeof jsonData !== 'object') {
    return;
  }

  const currentSheet = {};

  Object.keys(jsonData).forEach((key) => {
    const sheetData = jsonData[key];
    const sheetName = prefix + key;

    if (Array.isArray(sheetData) && sheetData.length > 0 && typeof sheetData[0] === 'object') {
      // Treat arrays of objects as separate sheets
      sheets[sheetName] = [];
      sheetData.forEach((item, index) => {
        const newSheetName = `${sheetName}_${index + 1}`;
        processJSON(item, sheets, newSheetName + '_');
        // Add processed data to the current sheet
        sheets[sheetName].push(...sheets[newSheetName]);
        delete sheets[newSheetName];
      });
    } else if (typeof sheetData === 'object') {
      // Treat nested objects as separate sheets
      processJSON(sheetData, sheets, sheetName + '_');
    } else {
      // Keep other variables in the same sheet
      currentSheet[key] = sheetData;
    }
  });

  if (Object.keys(currentSheet).length > 0) {
    sheets[prefix.slice(0, -1)] = [currentSheet];
  }
}

function writeToExcel(sheets, outputFileName) {
  const wb = XLSX.utils.book_new();

  Object.keys(sheets).forEach((sheetName) => {
    const ws = XLSX.utils.json_to_sheet(sheets[sheetName]);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  });

  XLSX.writeFile(wb, outputFileName);
}

function readAndProcessJSON(filePath, outputFileName) {
  try {
    const jsonData = JSON.parse(fs.readFileSync(filePath, 'utf8'));
    if (jsonData) {
      const sheets = {};
      processJSON(jsonData, sheets);
      writeToExcel(sheets, outputFileName);
      console.log(`Data successfully written to ${outputFileName}`);
    } else {
      console.error('Error: JSON data is null or undefined.');
    }
  } catch (error) {
    console.error('Error:', error.message);
  }
}

// Replace 'your_input_file.json' and 'output_file.xlsx' with your file names
readAndProcessJSON('nested.json', 'output_file.xlsx');
