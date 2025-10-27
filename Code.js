/**
 * Get arrest data for the bond application form
 * Returns properly formatted object with column headers as keys
 */
function getArrestDataForForm(bookingNumber, rowIndex) {
  Logger.log("=== getArrestDataForForm START ===");
  Logger.log("Booking Number: " + bookingNumber);
  Logger.log("Row Index: " + rowIndex);

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log("Spreadsheet: " + ss.getName());

    var sheet = ss.getSheetByName("Lee County Arrests");

    if (!sheet) {
      Logger.log("ERROR: Sheet not found, trying active sheet");
      sheet = ss.getActiveSheet();
    }

    Logger.log("Sheet name: " + sheet.getName());
    Logger.log("Sheet last row: " + sheet.getLastRow());
    Logger.log("Sheet last column: " + sheet.getLastColumn());

    // Get headers from row 1
    var lastCol = sheet.getLastColumn();
    var headerRange = sheet.getRange(1, 1, 1, lastCol);
    var headers = headerRange.getValues()[0];

    Logger.log("Headers count: " + headers.length);
    Logger.log("First 10 headers: " + headers.slice(0, 10).join(", "));

    // Validate row index
    if (rowIndex < 2 || rowIndex > sheet.getLastRow()) {
      throw new Error(
        "Invalid row index: " +
          rowIndex +
          ". Must be between 2 and " +
          sheet.getLastRow()
      );
    }

    // Get the data row
    var dataRange = sheet.getRange(rowIndex, 1, 1, lastCol);
    var dataRow = dataRange.getValues()[0];

    Logger.log("Data row length: " + dataRow.length);
    Logger.log("First 5 values: " + dataRow.slice(0, 5).join(", "));

    // Create the mapped object
    var formData = {};
    var mappedCount = 0;

    for (var i = 0; i < headers.length; i++) {
      var header = headers[i];
      var value = dataRow[i];

      // Only map non-empty headers
      if (header && header.toString().trim() !== "") {
        formData[header] = value;
        mappedCount++;
      }
    }

    Logger.log("Mapped " + mappedCount + " fields");
    Logger.log(
      "Form data keys: " + Object.keys(formData).slice(0, 10).join(", ")
    );

    // Log specific fields we care about
    Logger.log("Full_Name: " + formData["Full_Name"]);
    Logger.log("DOB: " + formData["DOB"]);
    Logger.log("Booking_Number: " + formData["Booking_Number"]);
    Logger.log("All_Charges: " + formData["All_Charges"]);
    Logger.log("Bond_Amount: " + formData["Bond_Amount"]);

    Logger.log("=== getArrestDataForForm SUCCESS ===");
    Logger.log(
      "Returning object with " + Object.keys(formData).length + " keys"
    );

    return formData;
  } catch (error) {
    Logger.log("=== getArrestDataForForm ERROR ===");
    Logger.log("Error type: " + error.name);
    Logger.log("Error message: " + error.message);
    Logger.log("Error stack: " + error.stack);

    throw new Error("Failed to retrieve arrest data: " + error.message);
  }
}

/**
 * Submit bond application form data
 */
function submitBondApplication(formData) {
  Logger.log("=== submitBondApplication START ===");

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Bond Applications");

    // Create sheet if it doesn't exist
    if (!sheet) {
      Logger.log("Creating Bond Applications sheet");
      sheet = ss.insertSheet("Bond Applications");

      // Add headers
      var headers = [
        "Timestamp",
        "Booking Number",
        "Defendant Full Name",
        "Defendant DOB",
        "Defendant Phone",
        "Defendant Email",
        "Defendant Address",
        "Defendant City",
        "Defendant State",
        "Defendant ZIP",
        "Charges",
        "Bond Amount",
        "Bond Type",
        "Case Number",
        "County",
        "Court Date",
        "Court Time",
        "Court Location",
        "Indemnitor Name",
        "Indemnitor Relationship",
        "Indemnitor Phone",
        "Indemnitor Email",
        "Indemnitor Address",
        "Indemnitor City",
        "Indemnitor State",
        "Indemnitor ZIP",
        "Indemnitor Employer",
        "Additional Notes",
      ];

      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    // Append form data
    var timestamp = new Date();
    var row = [
      timestamp,
      formData.bookingNumber || "",
      formData.defendantFullName || "",
      formData.defendantDOB || "",
      formData.defendantPhone || "",
      formData.defendantEmail || "",
      formData.defendantAddress || "",
      formData.defendantCity || "",
      formData.defendantState || "",
      formData.defendantZip || "",
      formData.charges || "",
      formData.bondAmount || "",
      formData.bondType || "",
      formData.caseNumber || "",
      formData.county || "",
      formData.courtDate || "",
      formData.courtTime || "",
      formData.courtLocation || "",
      formData.indemnitorName || "",
      formData.indemnitorRelationship || "",
      formData.indemnitorPhone || "",
      formData.indemnitorEmail || "",
      formData.indemnitorAddress || "",
      formData.indemnitorCity || "",
      formData.indemnitorState || "",
      formData.indemnitorZip || "",
      formData.indemnitorEmployer || "",
      formData.additionalNotes || "",
    ];

    sheet.appendRow(row);

    Logger.log("=== submitBondApplication SUCCESS ===");

    return {
      success: true,
      message: "Application submitted successfully",
      timestamp: timestamp,
    };
  } catch (error) {
    Logger.log("=== submitBondApplication ERROR ===");
    Logger.log("Error: " + error.message);

    throw new Error("Failed to submit application: " + error.message);
  }
}

/**
 * Test function - Run this in Apps Script to verify everything works
 */
function testGetArrestData() {
  Logger.log("=== TEST START ===");

  // CHANGE THESE TO MATCH YOUR ACTUAL DATA
  var bookingNumber = "1013788"; // Replace with actual booking number
  var rowIndex = 69; // Replace with actual row number (NOT zero-indexed)

  Logger.log("Testing with booking: " + bookingNumber + ", row: " + rowIndex);

  try {
    var data = getArrestDataForForm(bookingNumber, rowIndex);

    Logger.log("=== TEST RESULTS ===");
    Logger.log("Data type: " + typeof data);
    Logger.log("Is Array: " + Array.isArray(data));
    Logger.log("Keys count: " + Object.keys(data).length);
    Logger.log("All keys: " + Object.keys(data).join(", "));

    Logger.log("\n=== SAMPLE VALUES ===");
    Logger.log("Full_Name: " + data.Full_Name);
    Logger.log("DOB: " + data.DOB);
    Logger.log("Booking_Number: " + data.Booking_Number);
    Logger.log("All_Charges: " + data.All_Charges);
    Logger.log("Bond_Amount: " + data.Bond_Amount);
    Logger.log("Address: " + data.Address);
    Logger.log("City: " + data.City);
    Logger.log("State: " + data.State);
    Logger.log("ZIP: " + data.ZIP);

    Logger.log("\n=== TEST SUCCESS ===");
  } catch (error) {
    Logger.log("=== TEST FAILED ===");
    Logger.log("Error: " + error.message);
    Logger.log("Stack: " + error.stack);
  }
}

/**
 * Helper function to check sheet structure
 */
function debugSheetStructure() {
  Logger.log("=== SHEET STRUCTURE DEBUG ===");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Lee County Arrests");

  if (!sheet) {
    Logger.log('ERROR: Sheet "Lee County Arrests" not found');
    Logger.log("Available sheets:");
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      Logger.log("  - " + sheets[i].getName());
    }
    return;
  }

  Logger.log("Sheet found: " + sheet.getName());
  Logger.log("Last row: " + sheet.getLastRow());
  Logger.log("Last column: " + sheet.getLastColumn());

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log("Headers (" + headers.length + " total):");

  for (var i = 0; i < headers.length; i++) {
    Logger.log("  [" + i + "] " + headers[i]);
  }

  Logger.log("=== DEBUG COMPLETE ===");
}
