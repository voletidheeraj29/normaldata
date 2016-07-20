var express = require('express');
var app = express();

//Middleware
app.use('/user/:id', function (req, res, next) {
  console.log('Request Type:', req.method);
  //next();
});

//Middleware
app.use(function (req, res, next) {
  console.log('Time:', Date.now());
  next();
});

//Routing
app.get('/', function (req, res) {
  console.log('Hello World!');
  res.send('Hello World!');
});

app.get('/user/:id', function (req, res) {
  console.log('GET Request for user id: ', req.params.id);
  res.send('GET Request for user id: ' + req.params.id);
});

app.post('/user/:id', function (req, res) {
  console.log('POST Request for user id: ', req.params.id);
  res.send('POST Request for user id: ' + req.params.id);
});

app.listen(3000, function () {
  console.log('Example app listening on port 3000!');
});