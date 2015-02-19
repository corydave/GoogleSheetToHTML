/*
 * NOTE: The following variables need to be modified with customizeable data:
 *
 *  "LOGO" - The URL where a logo resides (that will appear on some user interfaces).
 *           Be aware that an image on Google Drive will not work.
 *
 *  "ACTUALURL" - The URL where users will be directed to go in order to see the HTML version.
 *
 *  "HTMLFILE" - The document key in Google Drive where the HTML file is (this key is used to open the file and 
 *               overwrite the content). Note that it MUST be in a folder that is open to the public.
 */
 

var LOGO = 'https://bbsupport.sln.suny.edu/bbcswebdav/institution/OSCQR.png';
var ACTUALURL = '';
var HTMLFILE = '';

/*
 * NOTE: There are multiple tabs that must exist for this code to work:
 *
 *  "Chronological" - This tab is where the content goes (it can be formatted with blank lines, colors, borders, etc.).
 *
 *  "forHTML" - This tab is where the content for the HTML page will be drawn from. Content from "Chronological" will be parsed
 *              and and placed here. 
 *
 *  "Standards" - This tab is where a nice, readable, formatted version of the content from "Chronological" resides.
 */
 


/**
 * @description Runs when the Spreadsheet is open - this function generates the OSCQR menu and hides the "Standards" sheet
 * @type function
 */
function onOpen() {  
 
  SpreadsheetApp.getUi()
    .createMenu('OSCQR')
    .addItem('Publish To Web', 'publish')
    .addItem('Generate Standards', 'standards')
    .addToUi();
    
  try {
    SpreadsheetApp.getActive().getSheetByName('Standards').hideSheet();
    SpreadsheetApp.getActive().getSheetByName('ForHTML').hideSheet();
  } catch (e) {
  
  }

}




/**
 * @description Clears and formats "ForHTML", runs the function to rewrite the HTML file ("index.html"), and shows confirmation
 * @type function
 */
function publish() {
  
  formatForHTMLSheet();
  createHTML();
  showConfirm();
  
}




/**
 * @description Calls clearAll() and imports select data to "ForHTML" from "Chronological". Note that the first 6 rows of "Chronological" are not imported as they are headers
 * @type function
 */
function formatForHTMLSheet() {
  // DELEGATES - clearAll()
  
  clearAll();
  
  try {  
  
    var ss = SpreadsheetApp.getActive();
    var chrono = ss.getSheetByName('Chronological');
    var web = ss.getSheetByName('ForHTML');
    
    web.getRange('A:A').setValues(chrono.getRange('E6:E').getValues());
    web.getRange('B:B').setValues(chrono.getRange('K6:K').getValues());
    web.getRange('C:C').setValues(chrono.getRange('M6:M').getValues());
    web.getRange('D:D').setValues(chrono.getRange('O6:O').getValues());
    
    removeSpaces('ForHTML');
    
  } catch (e) { 
  
    Browser.msgBox('Error in formatForHTMLSheet(), line ' + e.lineNumber);
    
  }
  
}




/**
 * @description Clears the "ForHTML" sheet and preps it to receive the data from "Chronological"
 * @type function
 */
function clearAll() {
  // HELPER - formatForHTMLSheet()
  
  try {
  
    var ss = SpreadsheetApp.getActive();
    var web = ss.getSheetByName('ForHTML');
    
    var numRows = ss.getSheetByName('Chronological').getMaxRows();
    web.clear();
      
    if (web.getMaxColumns() > 1) {
      web.deleteColumns(1, web.getMaxColumns()-1);
    }
    
    if (web.getMaxRows() > 1) {
      web.deleteRows(1, web.getMaxRows()-1);
    }
  
    web.insertRows(1, numRows-6);
    
  } catch (e) {
  
    Browser.msgBox('Error in clearAll(), line ' + e.lineNumber);
    
  }
  
}



/**
 * @description Walks through the rows in "ForHTML" and deletes any row with spaces (assuming that if the cell in column A is empty, there is no data)
 * @type function
 * @param {string} sheetName - The name of the sheet being passed in
 */
function removeSpaces(sheetName) {

  try {
  
    var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
     
    if (sheet.getMaxRows() > 2) {    
      for (var i = 1; i < sheet.getMaxRows(); i++) {
        if (sheet.getRange('A' + i).getValue() == '') {
          sheet.deleteRow(i);
        }
      }
    }
    
    if (sheet.getRange('A' + sheet.getMaxRows()).getValue() == '') {
      sheet.deleteRow(sheet.getMaxRows());
    }
    
  } catch (e) {
  
    Browser.msgBox('Error in removeSpaces(), line ' + e.lineNumber + '\\n' + e);
    
  }
  
}




/**
 * @description Generates a UiInstance with the OSCQR logo, a confirmation message, and a hyperlink to "index.html"
 * @type function
 */
function showConfirm() {

  try {
  
    var app = UiApp.createApplication();
    var flow = app.createFlowPanel(); // To simulate text + link
    
    flow.add(app.createImage(LOGO));
    flow.add(app.createHTML('<br /><br />'));
    flow.add(app.createInlineLabel('The annotations have been published succesfully.'));
    flow.add(app.createHTML('<br />'));
    flow.add(app.createAnchor('To see them, click here', ACTUALURL));
    
    flow.add(app.createInlineLabel('.'));
    flow.add(app.createHTML('<br /><br />'));  
    app.add(flow);
    
    var okButton = app.createButton('OK');
    okButton.setStyleAttribute('background', 'dodgerblue').setStyleAttribute('color', 'white');
    okButton.setWidth('100');
    okButton.setFocus(true);
    
    flow.add(okButton);  
    var ok = app.createServerHandler('ok');
    okButton.addClickHandler(ok);
    
    var doc = SpreadsheetApp.getActive();
    doc.show(app);
    
  } catch (e) {
  
    Browser.msgBox('showConfirm(), line ' + e.lineNumber);
    
  }
  
}




/**
 * @description Event handler to close the "publishing confirmation" box
 * @type function
 * @param {event} e - Button Event
 * @return app.close()
 */
function ok(e) {

  try {
  
    var app = UiApp.getActiveApplication();
    
    return app.close();
    
  } catch (error) {
  
    Browser.msgBox('Error in ok(), line ' + error.lineNumber);
    
  }
  
}




/**
 * @description Using a suave, sophisticated system of writing lines, taking content from cells, and implementing HTML, this function creates the HTML string for "index.html" and rewrites the file. Note that this method basically constructs the framework of "index.html", and invokes generateAnnotations() to parse the "ForHTML" sheet.
 * @type function
 */
function createHTML() {

  try {
    
    var htmlString = "<center><p class=\"oscqr\">OSCQR Annotations</p></center>\n\n" + generateAnnotations();
    
    var css = "<style>\n\n\n" + 
    "p.oscqr {\n  color: black;\n  font-family: Arial, Geneva, sans-serif;\n  font-size: 1.5em;\n  font-weight: bold;\n}\n\n" + 
    "table, td, tr, th {\n  border: 0px solid black;\n  padding: 15px;\n}\n\n" + 
    "p.standardTitle {\n  padding-left: 40px;\n  text-indent: -20px;\n  color: black;\n  font-size: 1.175em; /* 30px/16=1.875em */\n  font-family: Arial, Geneva, sans-serif;\n  font-weight: bold;\n}\n\n" + 
    "p.subHeading {\n  padding-left: 40px;\n  color: black;\n  font-family: Arial, Geneva, sans-serif;\n  font-size: 1em;\n  font-weight: bold;}\n\n" + 
    "p.standardBody {\n  padding-left: 60px;\n  color: black;\n  font-family: Arial, Geneva, sans-serif;\n}\n\n" +
    "body {\n  color: black;\n  font-size: 1em;\n  font-family: Arial, Geneva, sans-serif;\n}\n\n" + 
    "</style>\n\n";
    
    var linky = '<script src="https://rawgit.com/AnSavvides/jquery.linky/master/jquery.linky.min.js"></script>';
    var jquery = '<script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.3/jquery.min.js"></script>';
    var linkyCall = '<script type="text/javascript">$(document).ready(function(){$("body").linky();});</script>';
    
    var HTMLToPublish = "<HTML>\n<HEAD>\n<TITLE>OSCQR Annotations</TITLE>\n\n" + css + "\n" + jquery + "\n" + linky + "\n" + linkyCall + "\n\n</HEAD>\n\n<BODY>\n\n" + htmlString + "\n\n</BODY>\n</HTML>";
    
    var fileToModify = DriveApp.getFileById(HTMLFILE);

    fileToModify.setContent(HTMLToPublish);
    
  } catch (e) {
  
    Browser.msgBox('Error in createHTML(), line ' + e.lineNumber);  
    
  }

}




/**
 * @description Where the magic happens that generates the HTML for "index.html" based on the content in "ForHTML"
 * @type function
 * @return htmlString - a String that represents all the formatted standards in HTML 
 */
function generateAnnotations() {
  
  var htmlString = '';
  
  try {
  
    var web = SpreadsheetApp.getActive().getSheetByName('ForHTML');
    
    var _sanitize = function (text) {
      
      var returnedValue = text;
      
      // This code will clean up unicode characters that render well in the Spreadsheet, but display oddly when the HTML is served.
      // Note that these lines of code will grow as more offending unicode is discovered.
      returnedValue = returnedValue.replace(/\n/g, '<br />');
      returnedValue = returnedValue.replace(/\u2013/g, '-');
      returnedValue = returnedValue.replace(/\u2014/g, '-');
      returnedValue = returnedValue.replace(/\u201C/g, '"');
      returnedValue = returnedValue.replace(/\u201D/g, '"');
      returnedValue = returnedValue.replace(/\u2026/g, '...');
      returnedValue = returnedValue.replace(/\u00A0/g, '');
      // End sanitization
      
      return returnedValue;
      
      };
    
    for (var i = 1; i <= web.getLastRow(); i++) {
      htmlString += "\n<br /><br />\n\n<a name=\"" + _sanitize(web.getRange("A" + i).getValue()) + "\">" + "<p class=\"standardTitle\">" + web.getRange("A" + i).getValue() + ")&nbsp;" + web.getRange("B" + i).getValue() + "</p>\n\n";
      htmlString += "<p class=\"subHeading\">Annotated Explanations:</p>\n\n";
      htmlString += "<p class=\"standardBody\">" + _sanitize(web.getRange("C" + i).getValue()) + "</p>\n\n";
      htmlString += "<p class=\"subHeading\">Refresh Resources:</p>\n\n";
      htmlString += "<p class=\"standardBody\">" + _sanitize(web.getRange("D" + i).getValue()) + "</p>\n\n";
    }
    
    htmlString += "<br /><br /><br /><br />";
    
  } catch (e) {
  
    Browser.msgBox("Error in generateAnnotations(), line " + e.lineNumber);
    
  }
  
  return htmlString;
  
}




/**
 * @description Goes through "Chronological" and writes the standard number and standard text in each cell
 * @type function
 */
function standards() {
  
  try {  
  
    var ss = SpreadsheetApp.getActive();
    var stand = ss.getSheetByName('Standards');
    var chrono = ss.getSheetByName('Chronological');
    
    stand.showSheet();
    
    var numRows = ss.getSheetByName('Chronological').getMaxRows();
    stand.clear();
      
    if (stand.getMaxColumns() > 1) {
      stand.deleteColumns(1, stand.getMaxColumns()-1);
    }
    
    if (stand.getMaxRows() > 1) {
      stand.deleteRows(1, stand.getMaxRows()-1);
    }
  
    stand.insertRows(1, numRows-6);

    stand.getRange('A:A').setValues(chrono.getRange('E6:E').getValues());
    stand.getRange('B:B').setValues(chrono.getRange('K6:K').getValues());
    
    removeSpaces('Standards');
    stand.setColumnWidth(1, 33);
    stand.setColumnWidth(2, 800);
    formatAsTable('Standards');
    stand.getRange('A:B').setWrap(true);
    
    ss.setActiveSheet(stand);
    
  } catch (e) { 
  
    Browser.msgBox('Error in standards(), line ' + e.lineNumber + '\\n' + e);
    
  }
  
}




/**
 * @description Paints every other row gray and white, starting with white
 * @type function
 * @param {string} sheetName - The name of the sheet being passed in
 */
function formatAsTable(sheetName) {

  try {

  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);

    for (var i = 1; i <= sheet.getMaxRows(); i++) {
      if (i % 2 > 0) {
          sheet.getRange(i + ":" + i).setBackground('#F0F0F0').setVerticalAlignment('middle');
      } else {
          sheet.getRange(i + ":" + i).setBackground('#CCCCCC').setVerticalAlignment('middle');
      }
    }
  
  } catch (e) {
  
    Browser.msgBox('Error in formatAsTable(), line ' + e.lineNumber + '\\n' + e);
  
  }

}
