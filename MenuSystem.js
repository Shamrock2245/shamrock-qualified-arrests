function onOpen( ) {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🍀 Shamrock Automation')
    .addItem('📝 Open Bond Form (Selected Row)', 'openBondFormForSelectedRow')
    .addSeparator()
    .addItem('🔍 Generate Search Links (Selected Row)', 'generateSearchLinksForSelectedRow')
    .addItem('🌐 Open All Search Links (Selected Row)', 'openAllSearchLinks')
    .addToUi();
  
  Logger.log('✅ Qualified Arrests menu created');
}

function onInstall(e) {
  onOpen(e);
}

function openBondFormForSelectedRow() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var selectedRow = sheet.getActiveRange().getRow();
  
  if (selectedRow < 2) {
    ui.alert('Please select a data row (not the header row).');
    return;
  }
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var bookingCol = headers.indexOf('Booking_Number') + 1;
  
  if (bookingCol === 0) {
    ui.alert('Booking_Number column not found in spreadsheet.');
    return;
  }
  
  var bookingNumber = sheet.getRange(selectedRow, bookingCol).getValue();
  
  if (!bookingNumber) {
    ui.alert('Selected row has no booking number.');
    return;
  }
  
  Logger.log('Opening bond form for booking number: ' + bookingNumber + ' (row ' + selectedRow + ')');
  
  var template = HtmlService.createTemplateFromFile('Form');
  
  var htmlOutput = template.evaluate()
    .setWidth(1200)
    .setHeight(900)
    .setTitle('Shamrock Bail Bonds - Bond Application Form')
    .append('<script>window.FORM_DATA = {booking: "' + bookingNumber + '", row: ' + selectedRow + '};</script>');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Bond Application Form');
}

function generateSearchLinksForSelectedRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var selectedRow = sheet.getActiveRange().getRow();
  
  if (selectedRow < 2) {
    SpreadsheetApp.getUi().alert('Please select a data row.');
    return;
  }
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var searchLinksCol = headers.indexOf('Search_Links') + 1;
  var fullNameCol = headers.indexOf('Full_Name') + 1;
  
  if (searchLinksCol === 0 || fullNameCol === 0) {
    SpreadsheetApp.getUi().alert('Required columns not found.');
    return;
  }
  
  var fullName = sheet.getRange(selectedRow, fullNameCol).getValue();
  
  if (!fullName) {
    SpreadsheetApp.getUi().alert('No name found in selected row.');
    return;
  }
  
  var searchLinks = generateSearchLinksForName_(fullName);
  sheet.getRange(selectedRow, searchLinksCol).setValue(searchLinks);
  
  SpreadsheetApp.getUi().alert('✅ Search links generated!');
}

function openAllSearchLinks() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var selectedRow = sheet.getActiveRange().getRow();
  
  if (selectedRow < 2) {
    SpreadsheetApp.getUi().alert('Please select a data row.');
    return;
  }
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var searchLinksCol = headers.indexOf('Search_Links') + 1;
  
  if (searchLinksCol === 0) {
    SpreadsheetApp.getUi().alert('Search_Links column not found.');
    return;
  }
  
  var searchLinks = sheet.getRange(selectedRow, searchLinksCol).getValue();
  
  if (!searchLinks) {
    SpreadsheetApp.getUi().alert('No search links found. Generate them first.');
    return;
  }
  
  var links = searchLinks.split('\n');
  var html = '<html><body><script>';
  
  links.forEach(function(link) {
    var match = link.match(/https?:\/\/[^\s )]+/);
    if (match) {
      html += 'window.open("' + match[0] + '", "_blank");';
    }
  });
  
  html += 'google.script.host.close();</script></body></html>';
  
  var htmlOutput = HtmlService.createHtmlOutput(html).setWidth(100).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Opening Links...');
}

function generateSearchLinksForName_(fullName) {
  var encodedName = encodeURIComponent(fullName);
  
  var links = [];
  links.push('Google: https://www.google.com/search?q=' + encodedName );
  links.push('Facebook: https://www.facebook.com/search/top?q=' + encodedName );
  links.push('LinkedIn: https://www.linkedin.com/search/results/all/?keywords=' + encodedName );
  links.push('Instagram: https://www.instagram.com/explore/tags/' + encodedName.replace(/\s+/g, '' ));
  links.push('Twitter: https://twitter.com/search?q=' + encodedName );
  
  return links.join('\n');
}
