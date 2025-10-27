// Shamrock Bail Bonds - Google Apps Script
// Fixed version to work with Bond Application Form

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Shamrock Bail Bonds")
    .addItem("Open Bond Application Form", "openBondForm")
    .addItem("Generate PDF", "generatePDF")
    .addToUi();
}

/**
 * Opens the bond application form in a dialog
 */
function openBondForm() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var row = sheet.getActiveRange().getRow();

    if (row <= 1) {
      SpreadsheetApp.getUi().alert("Please select an arrest record row first.");
      return;
    }

    // Store the active row in script properties so the form can access it
    PropertiesService.getScriptProperties().setProperty("ACTIVE_ROW", row);

    var html = HtmlService.createHtmlOutputFromFile("Form")
      .setWidth(900)
      .setHeight(700)
      .setTitle("Bond Application Form");

    SpreadsheetApp.getUi().showModalDialog(html, "Bond Application Form");
  } catch (error) {
    Logger.log("Error in openBondForm: " + error.toString());
    SpreadsheetApp.getUi().alert("Error opening form: " + error.message);
  }
}

/**
 * Gets arrest data for the currently selected row
 * Returns data as array of [columnName, value] pairs
 */
function getArrestData() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var row = PropertiesService.getScriptProperties().getProperty("ACTIVE_ROW");

    if (!row) {
      throw new Error("No row selected. Please select an arrest record first.");
    }

    row = parseInt(row);

    // Get headers (row 1)
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    // Get data for the selected row
    var rowData = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

    // Create array of [header, value] pairs
    var formData = [];
    for (var i = 0; i < headers.length; i++) {
      formData.push([headers[i], rowData[i]]);
    }

    Logger.log("Retrieved data for row " + row);
    Logger.log("Data: " + JSON.stringify(formData));

    return formData;
  } catch (error) {
    Logger.log("Error in getArrestData: " + error.toString());
    throw new Error("Failed to load arrest data: " + error.message);
  }
}

/**
 * Saves bond application data
 * @param {Object} formData - The form data to save
 */
function saveBondApplication(formData) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var row = PropertiesService.getScriptProperties().getProperty("ACTIVE_ROW");

    if (!row) {
      throw new Error("No row selected.");
    }

    row = parseInt(row);

    // Get headers to map form data to columns
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    // Map form fields to sheet columns
    var fieldMapping = {
      defendantName: ["Defendant Name", "Name", "Defendant", "Full Name"],
      dob: ["DOB", "Date of Birth", "Birth Date", "Birthdate"],
      phone: ["Phone", "Phone Number", "Telephone", "Cell", "Mobile"],
      email: ["Email", "E-mail", "Mail"],
      address: ["Address", "Street", "Location", "Residence"],
      arrestDate: ["Arrest Date", "Date of Arrest", "Arrested"],
      bookingNumber: ["Booking Number", "Booking #", "Booking", "Book Number"],
      charges: ["Charges", "Charge", "Offense", "Crime"],
      bondAmount: ["Bond Amount", "Bond", "Bail", "Bail Amount"],
      courtDate: ["Court Date", "Hearing Date", "Court", "Hearing"],
      indemnitorName: ["Indemnitor Name", "Indemnitor", "Cosigner"],
      relationship: ["Relationship", "Relation"],
      indemnitorPhone: ["Indemnitor Phone", "Cosigner Phone"],
      indemnitorEmail: ["Indemnitor Email", "Cosigner Email"],
      notes: ["Notes", "Comments", "Additional Info", "Remarks"],
    };

    // Update each field
    for (var formField in formData) {
      if (fieldMapping[formField]) {
        var possibleHeaders = fieldMapping[formField];

        // Find matching column
        for (var i = 0; i < headers.length; i++) {
          var headerLower = headers[i].toString().toLowerCase();

          for (var j = 0; j < possibleHeaders.length; j++) {
            if (headerLower.includes(possibleHeaders[j].toLowerCase())) {
              // Update the cell
              sheet.getRange(row, i + 1).setValue(formData[formField]);
              Logger.log(
                "Updated " + headers[i] + " with: " + formData[formField]
              );
              break;
            }
          }
        }
      }
    }

    // Add timestamp
    var timestampCol = findColumn(headers, [
      "Last Updated",
      "Updated",
      "Timestamp",
    ]);
    if (timestampCol > 0) {
      sheet.getRange(row, timestampCol).setValue(new Date());
    }

    Logger.log("Successfully saved bond application for row " + row);
    return { success: true, message: "Bond application saved successfully" };
  } catch (error) {
    Logger.log("Error in saveBondApplication: " + error.toString());
    throw new Error("Failed to save bond application: " + error.message);
  }
}

/**
 * Helper function to find a column by multiple possible names
 */
function findColumn(headers, searchTerms) {
  for (var i = 0; i < headers.length; i++) {
    var headerLower = headers[i].toString().toLowerCase();
    for (var j = 0; j < searchTerms.length; j++) {
      if (headerLower.includes(searchTerms[j].toLowerCase())) {
        return i + 1; // Return 1-based column index
      }
    }
  }
  return -1; // Not found
}

/**
 * Generates PDF from the bond application form
 */
function generatePDF() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var row = sheet.getActiveRange().getRow();

    if (row <= 1) {
      SpreadsheetApp.getUi().alert("Please select an arrest record row first.");
      return;
    }

    // Get the data for PDF generation
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var rowData = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

    // Create a new Google Doc for the PDF
    var docName = "Bond_Application_" + rowData[0] + "_" + new Date().getTime();
    var doc = DocumentApp.create(docName);
    var body = doc.getBody();

    // Add title
    body
      .appendParagraph("BOND APPLICATION FORM")
      .setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    body.appendHorizontalRule();

    // Add data
    for (var i = 0; i < headers.length; i++) {
      if (rowData[i]) {
        var para = body.appendParagraph(headers[i] + ": " + rowData[i]);
        para.setSpacingAfter(8);
      }
    }

    // Save and close the document
    doc.saveAndClose();

    // Get PDF blob
    var docFile = DriveApp.getFileById(doc.getId());
    var pdfBlob = docFile.getAs("application/pdf");

    // Save PDF to Drive
    var pdfFile = DriveApp.createFile(pdfBlob);
    pdfFile.setName(docName + ".pdf");

    // Delete the temporary doc
    docFile.setTrashed(true);

    SpreadsheetApp.getUi().alert(
      "PDF generated successfully!\n\nFile: " +
        pdfFile.getName() +
        "\nURL: " +
        pdfFile.getUrl()
    );

    return pdfFile.getUrl();
  } catch (error) {
    Logger.log("Error in generatePDF: " + error.toString());
    SpreadsheetApp.getUi().alert("Error generating PDF: " + error.message);
    throw error;
  }
}

/**
 * Test function to debug data retrieval
 */
function testGetArrestData() {
  try {
    // Set a test row
    PropertiesService.getScriptProperties().setProperty("ACTIVE_ROW", "2");

    var data = getArrestData();
    Logger.log("Test data: " + JSON.stringify(data));

    return data;
  } catch (error) {
    Logger.log("Test error: " + error.toString());
    throw error;
  }
}
