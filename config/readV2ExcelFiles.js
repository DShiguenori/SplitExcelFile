const Excel = require('exceljs');
const fs = require('fs');
require('dotenv').config();
const path = require('path');
const projectRoot = path.dirname(require.main.filename);
const listOfRequiredFiles = ['']

module.exports.start = async () => {
	console.log('---------------------- Starting ----------------------');

	// Folder configuration:
	let EXCEL_INPUT_FOLDER = path.join(projectRoot, process.env.EXCEL_FILES_V2);
	let EXCEL_OUTPUT_FOLDER = path.join(projectRoot, 'output_files');

	// Create the output folder if it does not exist
	if (!fs.existsSync(EXCEL_OUTPUT_FOLDER)) {
		fs.mkdirSync(EXCEL_OUTPUT_FOLDER, { recursive: true });
	}

	// Get all the required excel files
	let files = fs.readdirSync(EXCEL_INPUT_FOLDER);
	for(const file of files){
		console.log(file);
	}

	// Check if the required files are in the folder

	// For each "Atendimento" in the "cadExame" file, get the respective "audiometria" data and get the respective data from "paciente" and "empresa"
	let atendimentosExcelFile = `${EXCEL_INPUT_FOLDER}\\${file}`;
	// let workbookOutput = new Excel.Workbook();

	// for (let file of files) {
	// 	let excelPathFile = `${EXCEL_INPUT_FOLDER}\\${file}`;
	// 	console.log(`-> READING...   ${excelPathFile} `);

	// 	const excelInput = new Excel.Workbook();
	// 	await excelInput.xlsx.readFile(excelPathFile);

	// 	// Open the excel file, read each worksheet:
	// 	excelInput.eachSheet((inSheet, sheetId) => {
	// 		let rowCount = inSheet.rowCount;

	// 		let outSheet = workbookOutput.getWorksheet(inSheet.name);
	// 		if (!outSheet) {
	// 			outSheet = workbookOutput.addWorksheet(inSheet.name);
	// 		}

	// 		for (let rowNumber = 1; rowNumber <= rowCount; rowNumber++) {
	// 			const row = inSheet.getRow(rowNumber);
	// 			const newRow = outSheet.addRow(row.values);
	// 			newRow.commit();
	// 		}
	// 	});
	// }

	//const outputFilename = `${EXCEL_OUTPUT_FOLDER}\\joined_files.xlsx`;
	//await workbookOutput.xlsx.writeFile(outputFilename);
	//console.log(`JOINED FILES SAVED ${outputFilename}`);

	console.log('---------------------- END Reading ----------------------');
};
