var debug = true;

// Background color for `code` spans
// ----------------------------------
// Google Docs lightest grey: #f3f3f3
// Github uses: #f3f3f3
// Stackoverflow: #eff0f1

var bgcolor = "#eff0f1";

function log(msg) {
  if (debug) Logger.log(msg);
}

function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Format Code', 'doFormatCode')
      .addToUi();
}

function onInstall(e) {
  // stub
}

function getParagraphHeadingAtts(par) {
  var body = DocumentApp.getActiveDocument().getBody();
  return body.getHeadingAttributes(par.getHeading());
}

function doFormatCode() {
  var doc = DocumentApp.getActiveDocument();
  var selection = doc.getSelection();
  if (!selection) {
    throw 'Please select some text you\'d like formatted as `code`.';
  }
  var rangeElements = selection.getSelectedElements();
  var rangeElement = rangeElements[0];
  if (!rangeElement.isPartial() || rangeElements.length > 1) {
    throw 'Please select some text you\'d like formatted as `code`. Don\'t include hard line breaks or images.';
  }  
  var element = rangeElement.getElement();
  var text = element.asText();
  var startIndex = rangeElement.getStartOffset();
  var endIndex = rangeElement.getEndOffsetInclusive();
  
  var string = text.getText();

  var alreadyAsTrailingChar = false;

  var substring = string.substring(startIndex, endIndex+1); 
  log('substring: ' + substring);

  // Adjust startIndex to include any adjacent ` character
  if (startIndex > 0 && string.charAt(startIndex - 1) == '`') {
    --startIndex;
    log('Adjusted startIndex to: ' + startIndex);
  } 
  // Adjust endIndex to include any adjacent ` character
  if (string.charAt(endIndex + 1) == '`') {
    ++endIndex;
    log('Adjusted endIndex to: ' + endIndex);
  }
  

  var paraHeadingAtts = getParagraphHeadingAtts(element.getParent()); //.asParagraph() not needed
  var paraHeadingFontSize = paraHeadingAtts[DocumentApp.Attribute.FONT_SIZE];
  var paraHeadingFontFamily = paraHeadingAtts[DocumentApp.Attribute.FONT_FAMILY];

  // Make the font size a bit smaller than that of the parent paragraph
  var fontSize = Math.floor(0.9 * paraHeadingFontSize);

  text.deleteText(startIndex, endIndex);

  endIndex = startIndex + substring.length - 1;
  
  text.insertText(startIndex, '`' + substring + '`');
  
  text.setBackgroundColor(startIndex, endIndex+2, bgcolor);
  text.setFontSize(startIndex, endIndex+2, fontSize);

  text.setForegroundColor(startIndex, startIndex, bgcolor);
  text.setFontSize(startIndex, startIndex, paraHeadingFontSize);
  text.setFontFamily(startIndex, startIndex, paraHeadingFontFamily);

  text.setFontFamily(startIndex+1, endIndex+1, "Courier New");
  
  text.setForegroundColor(endIndex+2, endIndex+2, bgcolor);
  text.setFontSize(endIndex+2, endIndex+2, paraHeadingFontSize);
  text.setFontFamily(endIndex+2, endIndex+2, paraHeadingFontFamily);
  
  // Preserve selection  
  var rangeBuilder = doc.newRange();
  rangeBuilder.addElement(element, startIndex+1, endIndex+1);
  doc.setSelection(rangeBuilder.build());
}
