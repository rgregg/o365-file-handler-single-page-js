var express = require('express');
var app = express();
var bodyParser = require('body-parser');
var cookieParser = require('cookie-parser');

app.use(bodyParser.urlencoded({ extended: false }));
app.use(cookieParser());
app.use(express.static(__dirname + "/content"));

app.post('/filehandler/open', function (req, res) {
    console.log("Content-Type %s", req.headers["content-type"]);
    console.log("POST data: %s", JSON.stringify(req.body));

    switchToViewMode("open", res, req);
});

app.post('/filehandler/preview', function (req, res) {
    console.log("Content-Type %s", req.headers["content-type"]);
    console.log("POST data: %s", JSON.stringify(req.body));

    switchToViewMode("preview", req, res);
});

app.post('/filehandler/newfile', function(req, res) {
    console.log("Content-Type %s", req.headers["content-type"]);
    console.log("POST data: %s", JSON.stringify(req.body));

    switchToViewMode("newfile", res, req);
});

var server = app.listen(9999, function () {
    var host = server.address().address;
    var port = server.address().port;

    console.log("Example app listening at http://%s:%s", host, port);
});

function switchToViewMode(mode, req, res)
{
    console.log("Switching to view mode: %s", mode);
    
    // Convert the activation parameters into JSON
    var params = convertActivationParameters(req);
    console.log("activation params: %s", JSON.stringify(params));

    // Write the parameters out in the form of a cookie
    let options = {
      httpOnly: false  
    };

    res.cookie('fileHandlerActivation', JSON.stringify(params), options);
    res.writeHead(301, {
        Location: "http" + (req.socket.encrypted ? "s" : "") + "://" + req.headers["host"] + "/index.html?view=open"
    });
    res.end();
}

function convertActivationParameters(req) {
    var data = {
        "locale": req.body["cultureName"],
        "client": req.body["client"],
        "userId": req.body["userId"]
    }
    
    var encodedItems = req.body["items"];
    var items = JSON.parse(encodedItems);
    data.items = items;

    return data;
}