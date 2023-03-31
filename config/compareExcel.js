const Excel = require('exceljs');
const fs = require('fs');
require('dotenv').config();
const path = require('path');
const projectRoot = path.dirname(require.main.filename);

module.exports.start = async () => {
	console.log('---------------------- Starting Comparison ----------------------');

	const EXCEL_COMPARE_FOLDER = path.join(projectRoot, process.env.EXCEL_FILES_TO_COMPARE);

	const files = fs.readdirSync(EXCEL_COMPARE_FOLDER);
	if (files.length !== 2) {
		throw new Error('The number of files in the folder is not 2');
	}

	const file1 = path.join(EXCEL_COMPARE_FOLDER, files[0]);
	const file2 = path.join(EXCEL_COMPARE_FOLDER, files[1]);

	const workbook1 = new Excel.Workbook();
	await workbook1.xlsx.readFile(file1);

	const workbook2 = new Excel.Workbook();
	await workbook2.xlsx.readFile(file2);

	let differences = [];

	workbook1.eachSheet((sheet1, sheetId) => {
		const sheet2 = workbook2.getWorksheet(sheet1.name);

		if (!sheet2) {
			console.log(`Worksheet "${sheet1.name}" not found in ${file2}`);
			return;
		}

		const rowCount1 = sheet1.rowCount;
		const rowCount2 = sheet2.rowCount;

		if (rowCount1 !== rowCount2) {
			differences.push(
				`Worksheet "${sheet1.name}": Row count mismatch - ${rowCount1} rows in ${file1} and ${rowCount2} rows in ${file2}`,
			);
		}

		for (let rowNumber = 1; rowNumber <= Math.max(rowCount1, rowCount2); rowNumber++) {
			const row1 = sheet1.getRow(rowNumber);
			const row2 = sheet2.getRow(rowNumber);

			if (row1.getCell(1).value !== row2.getCell(1).value) {
				differences.push(
					`Worksheet "${sheet1.name}", Row ${rowNumber}: First cell value mismatch - ${
						row1.getCell(1).value
					} in ${file1} and ${row2.getCell(1).value} in ${file2}`,
				);
			}
		}
	});

	if (differences.length === 0) {
		console.log('No differences found');
	} else {
		console.log('Differences found:');
		for (const difference of differences) {
			console.log(difference);
		}
	}

	console.log('---------------------- End Comparison ----------------------');
};
