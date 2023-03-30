const express = require('express');
const bodyParser = require('body-parser');
const splitExel = require('./splitExel');

module.exports.initExcel = function () {
	splitExel.start();
};

module.exports.init = function () {
	let app = express();

	this.initExcel();

	return app;
};
