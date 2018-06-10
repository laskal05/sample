const async = require ('async');
const officegen = require('officegen');
const fs = require('fs');
const path = require('path');

var format = function() {
  return new Promise(function(resolve, reject) {
    var themeXml = fs.readFileSync ( path.resolve ( __dirname, 'themes/testTheme.xml' ), 'utf8' );
    
    var docx = officegen({
      type: 'docx',
      orientation: 'portrait',
      pageMargins: { top: 1000, left: 1000, bottom: 1000, right: 1000 }
    });
    
    docx.on( 'error', function ( err ) {
      console.log ( err );
    });
    
    var pObj = docx.createP ();
    
    pObj.addText('Simple');
    pObj.addText(' with color', { color: '000088' } );
    pObj.addText(' and back color.', { color: '00ffff', back: '000088' } );
    
    var pObj = docx.createP ();
    
    pObj.addText('Since ');
    pObj.addText('officegen 0.2.12', { back: '00ffff', shdType: 'pct12', shdColor: 'ff0000' } ); // Use pattern in the background.
    pObj.addText(' you can do ' );
    pObj.addText('more cool ', { highlight: true } ); // Highlight!
    pObj.addText('stuff!', { highlight: 'darkGreen' } ); // Different highlight color.
    
    var pObj = docx.createP();
    
    pObj.addText ('Even add ');
    pObj.addText ('external link', { link: 'https://github.com' });
    pObj.addText ('!');
    
    var pObj = docx.createP ();
    
    pObj.addText ( 'Bold + underline', { bold: true, underline: true } );
    
    var pObj = docx.createP ( { align: 'center' } );
    
    pObj.addText ( 'Center this text', { border: 'dotted', borderSize: 12, borderColor: '88CCFF' } );
    
    var pObj = docx.createP ();
    pObj.options.align = 'right';
    
    pObj.addText ( 'Align this text to the right.' );
    
    var pObj = docx.createP ();
    
    pObj.addText ( 'Those two lines are in the same paragraph,' );
    pObj.addLineBreak ();
    pObj.addText ( 'but they are separated by a line break.' );
    
    docx.putPageBreak ();
    
    var pObj = docx.createP ();
    
    pObj.addText ( 'Fonts face only.', { font_face: 'Arial' } );
    pObj.addText ( ' Fonts face and size.', { font_face: 'Arial', font_size: 40 } );
    
    var out = fs.createWriteStream ( 'tmp/out.docx' );
    
    out.on ( 'error', function ( err ) {
      console.log ( err );
    });
    
    async.parallel ([
      function ( done ) {
        out.on ( 'close', function () {
        console.log ( 'Finish to create a DOCX file.' );
        done ( null );
        });
        docx.generate ( out );
      }
    ], function ( err ) {
      if ( err ) {
        console.log ( 'error: ' + err );
      } // Endif.
      resolve();
    });
  });
};

var express = require('express');
var app = express();

var docxPath = '/tmp/out.docx';

//HelloWorld
app.get('/', function(req,res,next) {
  format().then(function() {
    res.send('<a href="' + docxPath + '">Download docx</a>');
  });
});

app.get('/tmp/out.docx', function(req,res,next) {
  // res.download(path.resolve(__dirname, '/tmp/out.docx'));
  res.sendFile(path.resolve(__dirname, 'tmp/out.docx'));
});

//Listen
app.listen(3000, function(){
  console.log("Start Express on port 3000.");
});
