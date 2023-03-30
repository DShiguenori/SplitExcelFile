const Excel = require('exceljs');
const fs = require('fs');
require('dotenv').config();
const path = require('path');
const projectRoot = path.dirname(require.main.filename);

module.exports.start = async () => {
	console.log('---------------------- Starting ----------------------');

	// Folder configuration:
	let EXCEL_INPUT_FOLDER = path.join(projectRoot, process.env.EXCEL_INPUT_FOLDER);
	let EXCEL_OUTPUT_FOLDER = path.join(projectRoot, process.env.EXCEL_OUTPUT_FOLDER);

	let NUM_ROWS_EACH_FILE = parseInt(process.env.NUMBER_ROWS_OUTPUT);

	let files = fs.readdirSync(EXCEL_INPUT_FOLDER);

	// Leitura para cada Arquivo (filial)
	for (let file of files) {
		let excelPathFile = `${EXCEL_INPUT_FOLDER}\\${file}`;
		console.log(`-> READING...   ${excelPathFile} `);

		const excelInput = new Excel.Workbook();
		await excelInput.xlsx.readFile(excelPathFile);

		// Open the excel file, read each worksheet:
		excelInput.eachSheet((inSheet, sheetId) => {
			let rowCount = inSheet.rowCount;

			(async () => {
				let currentFileCount = 1;
				let currentRowInOutput = 1;
				let workbookOutput = new Excel.Workbook();
				let outSheet = workbookOutput.addWorksheet(inSheet.name);

				for (let rowNumber = 1; rowNumber <= rowCount; rowNumber++) {
					const row = inSheet.getRow(rowNumber);
					await copyRow(row, currentRowInOutput, outSheet);

					if (currentRowInOutput === NUM_ROWS_EACH_FILE || rowNumber === rowCount) {
						const outputFilename = `${EXCEL_OUTPUT_FOLDER}\\out_${file}_${currentFileCount}.xlsx`;
						await workbookOutput.xlsx.writeFile(outputFilename);
						console.log(`FILE SAVED ${outputFilename}`);

						currentFileCount++;
						currentRowInOutput = 1;

						workbookOutput = new Excel.Workbook();
						outSheet = workbookOutput.addWorksheet(inSheet.name);
					} else {
						currentRowInOutput++;
					}
				}
			})();
		});
	}
	console.log('---------------------- END READING FOLDERS AND THEIR FILES ----------------------');
};

async function copyRow(row, rowNumber, outSheet) {
	// Copy the row to the output sheet
	const newRow = outSheet.getRow(rowNumber);
	newRow.values = row.values;
	newRow.commit();
}

// Read a cell, outputs a string
function readCellString(row, number) {
	let rowValue = row.values[number];

	if (rowValue == null) return null;

	return rowValue.toString();
}
