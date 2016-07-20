const http = require('http');
var movies = require('./movies');

const hostname = '127.0.0.1';
const port = 1337;

http.createServer((req, res) => {
		console.log( ` request for URL `  + req.url);
		var toBeSent = "Hello World";
		if (req.url == "/movies/matrix") {
			toBeSent = movies.matrix();
		}
		res.writeHead(200, {
			'Content-Type' : 'text/plain'
		});
		res.end(toBeSent);
	}).listen(port, hostname, () => {
		console.log( ` Server running at http : //${hostname}:${port}/`);
		});
