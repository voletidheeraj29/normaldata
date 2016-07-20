var fs = require("fs");
function readFile(callback) {
	fs.readFile("./dataFile.txt", function (err, data) {
		callback(err, data);
	});
}
module.exports.data = readFile;
