const mammoth = require('mammoth');
const XLSX = require('xlsx');
const { JSDOM } = require('jsdom');

//#region MAIN
async function convertWordTablesToExcel(inputDoc, outputDoc) {
  try {
    //* Use mammoth to convert the input docx to html
    const result = await mammoth.convertToHtml({ path: inputDoc });

    //* Use JSDOM to extract tables from the html created above
    const tables = extractTablesFromHtml(result.value);

    //* Combine all tables into one array called combinedData
    const combinedData = [];

    tables.forEach((table, index) => {
      combinedData.push(...table);

      //* Separate tables with empty row
      if (index !== tables.length - 1) {
        combinedData.push([]);
      }
    });

    //* Use XLSX to create a new workbook and a sheet called 'AllData'
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet(combinedData);
    XLSX.utils.book_append_sheet(workbook, sheet, `AllData`);

    // tables.forEach((table, index) => {
    //   const sheet = XLSX.utils.aoa_to_sheet(table);
    //   XLSX.utils.book_append_sheet(workbook, sheet, `Table${index + 1}`);
    // });

    XLSX.writeFile(workbook, outputDoc);
    console.log('Success!');
  } catch (error) {
    console.error('Error: ', error);
  }
}
//#endregion MAIN

//* Use JSDOM to find all <tables> in the html output
//* return those to convertWordTablesToExcel function
function extractTablesFromHtml(html) {
  const tables = [];
  const dom = new JSDOM(html);
  const document = dom.window.document;
  const tableElements = document.querySelectorAll('table');
//   console.log('🚀 ~ file: process-word.js:29 ~ extractTablesFromHtml ~ tableElements:', tableElements);

  tableElements.forEach((tableElem, tableIndex) => {
    const rows = [];
    const rowElements = tableElem.querySelectorAll('tr');
    // console.log('🚀 ~ file: process-word.js:35 ~ tableElements.forEach ~ rowElements:', rowElements);

    rowElements.forEach((rowElem, rowIndex) => {
      const cells = [];
      const cellElements = rowElem.querySelectorAll('td, th');
    //   console.log('🚀 ~ file: process-word.js:39 ~ rowElements.forEach ~ cellElements:', cellElements);

      cellElements.forEach((cell) => {
        cells.push(cell.textContent.trim());
      });
      rows.push(cells);
    });
    tables.push(rows);
  });

  return tables;
}

//* Location of your input docx
//* Location of your output xlsx
const inputPath = 'example.docx';
const outputPath = 'example.xlsx';

convertWordTablesToExcel(inputPath, outputPath);
//! Run this from command line using `node word-2-excel.js`
