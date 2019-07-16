var xml2js = require('xml2js');
var when = require('when');
var _ = require('lodash');

var util = require('util');

function convert(xml) {

  var attrkey = 'attributes';
  var charkey = '$value'; // added this to get the correct key > value that exchange expects

  var parser = new xml2js.Parser({
    attrkey: attrkey,
    charkey: charkey,
    trim: true,
    ignoreAttrs: false,
    explicitRoot: false,
    explicitCharkey: false,
    explicitArray: false,
    explicitChildren: false,
    tagNameProcessors: [
      function(tag) {
        return tag.replace('t:', '').replace('m:',''); // and this to cleanup some extra tags,
                                                       // not specifically used in this example but it is needed
      }
    ]
  });

  return when.promise((resolve, reject) => {
    parser.parseString(xml, (err, result) => {
      if(err) reject(err);
      else {
        var ewsFunction = _.keys(result['soap:Body'])[0];
        var parsed = result['soap:Body'][ewsFunction];
        parsed[attrkey] = _.omit(parsed[attrkey], ['xmlns', 'xmlns:t']);
        if(_.isEmpty(parsed[attrkey])) parsed = _.omit(parsed, [attrkey]);
        resolve(parsed);
      }
    });
  });

}

//var xml = '<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"        xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"        xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"        xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">  <soap:Header>    <t:RequestServerVersion Version="Exchange2007_SP1" />  </soap:Header>  <soap:Body>    <m:FindItem Traversal="Shallow">      <m:ItemShape>        <t:BaseShape>IdOnly</t:BaseShape>        <t:AdditionalProperties>          <t:FieldURI FieldURI="item:Subject" />          <t:FieldURI FieldURI="calendar:Start" />          <t:FieldURI FieldURI="calendar:End" />        </t:AdditionalProperties>      </m:ItemShape>      <m:CalendarView MaxEntriesReturned="5" StartDate="2013-08-21T17:30:24.127Z" EndDate="2013-09-20T17:30:24.127Z" />      <m:ParentFolderIds>        <t:FolderId Id="AAMk" ChangeKey="AgAA" />      </m:ParentFolderIds>    </m:FindItem>  </soap:Body></soap:Envelope>';

/*
convert(xml).then(json => {
  console.log('ewsArgs = ' + util.inspect(json, false, null));
});
*/

//express stuff

var express = require('express');
var app = express();
var bodyParser = require('body-parser');

var urlencodedParser = bodyParser.urlencoded({ extended: false })


// This responds with "Hello World" on the homepage
app.get('/', function (req, res) {
   res.sendFile( __dirname + "/" + "index.html" );
})

// This responds with "Hello World" on the homepage
app.post('/convert', function (req, res) {
  data = req.query.data;
  //remove newlinews
  //data = data.replace(/\n/g, '');
  jsonData =  convert(data);
  convert(data).then(jsonData => {
    res.send(JSON.stringify(jsonData));
  }).catch(
    (err) => {
      res.send("error" + JSON.stringify(err));
    }
  );
  
})


var server = app.listen(8081, urlencodedParser, function () {
   var host = server.address().address
   var port = server.address().port
   
   console.log("Example app listening at http://%s:%s", host, port)
})