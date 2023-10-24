const mammoth = require('mammoth');
const XLSX = require('xlsx');
const { JSDOM } = require('jsdom');
const fs = require('fs');

//#region MAIN
async function convertWordTablesToJson(inputDoc, outputJson) {
  try {
    //* Use mammoth to convert the input docx to html
    const result = await mammoth.convertToHtml({ path: inputDoc });

    //* Use JSDOM to extract tables from the html created above
    const tables = extractTablesFromHtml(result.value);
    console.log("🚀 ~ file: word-2-json.js:14 ~ convertWordTablesToJson ~ tables:", tables)

    //* new translations object
    const translations = {};

    tables.forEach((table) => {
      table.forEach((row) => {
        if (row.length >= 2) {
          const key = row[0].trim();
          console.log("🚀 ~ file: word-2-json.js:23 ~ table.forEach ~ key:", key)
          const value = row[1].trim();
          console.log("🚀 ~ file: word-2-json.js:25 ~ table.forEach ~ value:", value)
          translations[key] = value;
        }
      });
    });

    fs.writeFileSync(outputJson, JSON.stringify(translations, null, 2));
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
//* Location of your output json
const inputPath = 'example.docx';
const outputPath = 'translation.json';

convertWordTablesToJson(inputPath, outputPath);
//! Run this from command line using `node word-2-excel.js`
