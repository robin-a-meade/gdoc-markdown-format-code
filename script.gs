var debug = false;

// Background color for `code` spans
// ----------------------------------
// Google Docs lightest grey: #f3f3f3
// Github uses: #f3f3f3
// Stackoverflow uses: #eff0f1

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
    DocumentApp.getUi().alert('Please select some text you\'d like formatted as `code`.');
    return;
  }
  var rangeElements = selection.getSelectedElements();
  var rangeElement = rangeElements[0];
  if (!rangeElement.isPartial() || rangeElements.length > 1) {
    DocumentApp.getUi().alert('Please select some text you\'d like formatted as `code`. Don\'t include hard line breaks or images.');
    return;
  }  
  var element = rangeElement.getElement();
  //DocumentApp.getUi().alert('Type: ' + element.getType());

  var text = element.asText();
  var startIndex = rangeElement.getStartOffset();
  var endIndex = rangeElement.getEndOffsetInclusive();
  
  var string = text.getText();

  var substring = string.substring(startIndex, endIndex+1); 
  
  var paraHeadingAtts = getParagraphHeadingAtts(element.getParent()); //.asParagraph() not needed
  var paraHeadingFontSize = paraHeadingAtts[DocumentApp.Attribute.FONT_SIZE];
  var paraHeadingFontFamily = paraHeadingAtts[DocumentApp.Attribute.FONT_FAMILY];
  
  var baseFontSize = null;
  // Determine font-size of left or right character
  if (startIndex > 0) baseFontSize = text.getFontSize(startIndex - 1);
  if (baseFontSize == null && endIndex < string.length - 1) baseFontSize = text.getFontSize(endIndex + 1);
  if (baseFontSize == null) baseFontSize = paraHeadingFontSize;
  
  // Make the monospace font size a little bit smaller than the baseFontSize
  var fontSize = Math.floor(0.9 * baseFontSize);

  // Adjust startIndex to include any adjacent ` character
  if (startIndex > 0 && string.charAt(startIndex - 1) == '`') {
    --startIndex;
  } 
  // Adjust endIndex to include any adjacent ` character
  if (string.charAt(endIndex + 1) == '`') {
    ++endIndex;
  }

  text.deleteText(startIndex, endIndex);

  endIndex = startIndex + substring.length - 1;
  
  text.insertText(startIndex, '`' + substring + '`');
  
  text.setBackgroundColor(startIndex, endIndex+2, bgcolor);
  text.setFontSize(startIndex, endIndex+2, fontSize);

  text.setForegroundColor(startIndex, startIndex, bgcolor);
  text.setFontSize(startIndex, startIndex, baseFontSize);
  text.setFontFamily(startIndex, startIndex, paraHeadingFontFamily);

  // Use "Courier New" as the monospace font
  text.setFontFamily(startIndex+1, endIndex+1, "Courier New");
  
  text.setForegroundColor(endIndex+2, endIndex+2, bgcolor);
  text.setFontSize(endIndex+2, endIndex+2, baseFontSize);
  text.setFontFamily(endIndex+2, endIndex+2, paraHeadingFontFamily);
  
  // Preserve selection  
  var rangeBuilder = doc.newRange();
  rangeBuilder.addElement(element, startIndex+1, endIndex+1);
  doc.setSelection(rangeBuilder.build());
}
