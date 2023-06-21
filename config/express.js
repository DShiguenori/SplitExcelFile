const express = require('express');
const bodyParser = require('body-parser');

module.exports.initExcel = function () {
	const argv = require('minimist')(process.argv.slice(2));
	const mode = argv.mode;

	const readV2ExcelFiles = require('./readV2ExcelFiles');
	readV2ExcelFiles.start();
};

module.exports.init = function () {
	let app = express();

	this.initExcel();

	return app;
};
