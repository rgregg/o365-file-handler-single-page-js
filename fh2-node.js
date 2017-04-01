var express = require('express');
var app = express();
var bodyParser = require('body-parser');
var cookieParser = require('cookie-parser');

app.use(bodyParser.urlencoded({ extended: false }));
app.use(cookieParser());
app.use(express.static(__dirname + "/content"));

app.post('/filehandler/open', function (req, res) {
    switchToViewMode("open", req, res);
});

app.post('/filehandler/preview', function (req, res) {
    switchToViewMode("preview", req, res);
});

app.post('/filehandler/newfile', function(req, res) {
    switchToViewMode("newfile", req, res);
});

var server = app.listen(9999, function () {
    var host = server.address().address;
    var port = server.address().port;

    console.log("Example app listening at http://%s:%s", host, port);
});

function switchToViewMode(mode, req, res)
{
    console.log("Switching to view mode: %s", mode);
    console.log("Content-Type %s", req.headers["content-type"]);
    console.log("POST data: %s", JSON.stringify(req.body));
   
    // Convert the activation parameters into JSON
    var params = convertActivationParameters(req.body);
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

function convertActivationParameters(postBody) {

    if (postBody == undefined)
    {
        console.log("no activation parameters to convert...");
        return null;
    }

    var data = {
        "locale": postBody["cultureName"],
        "client": postBody["client"],
        "userId": postBody["userId"],
        "items": null
    }
    
    var encodedItems = postBody["items"];
    var items = JSON.parse(encodedItems);
    data.items = items;

    return data;
}