function getArrestDataForForm(bookingNumber, rowIndex) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    var data = {};
    headers.forEach(function(header, index) {
      data[header] = rowData[index];
    });
    
    Logger.log('Retrieved arrest data for booking ' + bookingNumber);
    return data;
    
  } catch (error) {
    Logger.log('Error getting arrest data: ' + error.message);
    throw error;
  }
}
