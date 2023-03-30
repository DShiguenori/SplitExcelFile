const express = require('express');
const bodyParser = require('body-parser');

module.exports.initExcel = function () {
	const argv = require('minimist')(process.argv.slice(2));
	const mode = argv.mode;
	console.log(`Running in "${mode}" mode`);

	if (mode === 'join') {
		const joinExel = require('./joinExcel');
		joinExel.start();
	} else if (mode === 'split') {
		const splitExcel = require('./splitExel');
		splitExcel.start();
	}
};

module.exports.init = function () {
	let app = express();

	this.initExcel();

	return app;
};
