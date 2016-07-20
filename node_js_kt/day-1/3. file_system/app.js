const http = require('http');
var files = require('./files');

const hostname = '127.0.0.1';
const port = 1337;

http.createServer((req, res) => {
		console.log( ` request for URL `  + req.url);
		res.writeHead(200, {
			'Content-Type' : 'text/plain'
		});
		if (req.url == "/file-as-text") {
			files.data(function (err, data) {
				if(err){
					res.end("error");
				} else {
					console.log('should be printed after');
					res.end(data);
				}
			});
			console.log('should be printed before');
		} else {
			res.end("Hello World!!!");
		}
	}).listen(port, hostname, function () {
	console.log('Server running at http://' + hostname + ':' + port + '/');
});
