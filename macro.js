function createSheetsAndGenerateStyledPDF() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName('Sheet1');

  // Check if the sheet exists
  if (!sourceSheet) {
    Logger.log("Sheet 'Sheet1' not found");
    SpreadsheetApp.getUi().alert("Sheet 'Sheet1' not found. Please ensure the sheet exists.");
    return;
  }

  var data = sourceSheet.getDataRange().getValues();
  var createdSheets = [];

  // Initialize variables
  var currentDivision = '';
  var currentCustomer = '';
  var currentPalletNo = '';
  var totalPallets = 0;

  // Calculate total pallets for each division
  var divisionPalletCount = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0] && row[0].toUpperCase() !== 'TOTAL') {
      currentDivision = row[0];
      if (!divisionPalletCount[currentDivision]) {
        divisionPalletCount[currentDivision] = 1;
      } else {
        divisionPalletCount[currentDivision]++;
      }
    }
  }

  currentDivision = '';
  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    // Skip the "TOTAL" row
    if (row[0].toUpperCase() === 'TOTAL') {
      continue;
    }

    // Check for new division
    if (row[0] && row[0].toUpperCase() !== 'TOTAL') {
      currentDivision = row[0];
      currentPalletNo = row[1];
      currentCustomer = row[2];
      totalPallets = divisionPalletCount[currentDivision];
      var sheetName = currentDivision + " " + currentPalletNo + " OF " + totalPallets;
      
      // Check for invalid characters in sheet name
      sheetName = sheetName.replace(/[\/:*?"<>|]/g, '-');
      
      var sheet = ss.getSheetByName(sheetName);

      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        createdSheets.push(sheet);
        sheet.appendRow([currentDivision + " " + currentPalletNo + " OF " + totalPallets]);
        sheet.appendRow([currentCustomer]);
        // Skip adding the third row
        sheet.appendRow(['Product name', 'Quantity']);

        // Merge A1 and A2, style headers
        sheet.getRange('A1:B1').merge().setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle');
        sheet.getRange('A2:B2').merge().setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center').setVerticalAlignment('middle');
        sheet.getRange('A3:B3').setFontWeight('bold').setFontSize(12).setBackground('#f2f2f2').setBorder(true, true, true, true, true, true).setHorizontalAlignment('center').setVerticalAlignment('middle');

        // Center the table on the page
        sheet.setColumnWidths(1, 2, 300); // Increased column width to accommodate longer text
        sheet.setRowHeights(1, 3, 25);
      }
    }

    // Add product data with alternating row colors
    for (var j = 3; j < data[0].length; j++) {
      if (row[j] && !isNaN(row[j])) {
        var newRow = sheet.appendRow([data[0][j], row[j]]);
        var lastRow = sheet.getLastRow();
        var range = sheet.getRange('A' + lastRow + ':B' + lastRow);
        range.setBorder(true, true, true, true, true, true).setFontWeight('normal').setBackground(null).setHorizontalAlignment('center').setVerticalAlignment('middle');
        if (lastRow % 2 === 0) {
          range.setBackground('#f9f9f9'); // Light grey background for even rows
        } else {
          range.setBackground('#ffffff'); // White background for odd rows
        }
        sheet.getRange('A' + lastRow).setFontSize(10); // Ensure consistent font size for all products
      }
    }
  }

  // Create a temporary spreadsheet and copy created sheets to it
  var tempSpreadsheet = SpreadsheetApp.create('Temp Spreadsheet for PDF');
  var tempSsId = tempSpreadsheet.getId();
  var tempSs = SpreadsheetApp.openById(tempSsId);

  createdSheets.forEach(function(sheet) {
    sheet.copyTo(tempSs).setName(sheet.getName());
  });

  // Remove the default "Sheet1" in the temporary spreadsheet
  var tempSheet = tempSs.getSheetByName('Sheet1');
  if (tempSheet) {
    tempSs.deleteSheet(tempSheet);
  }

  // Export the temporary spreadsheet as PDF
  var url = 'https://docs.google.com/spreadsheets/d/' + tempSsId + '/export?';
  var url_ext = 'exportFormat=pdf&format=pdf' +
                '&size=A4' +
                '&portrait=true' +
                '&fitw=true' +
                '&sheetnames=false&printtitle=false' +
                '&pagenumbers=false&gridlines=false' +
                '&fzr=false';

  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url + url_ext, {
    headers: {
      'Authorization': 'Bearer ' + token
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() == 200) {
    var pdfBlob = response.getBlob().setName(ss.getName() + '.pdf');
    var base64data = Utilities.base64Encode(pdfBlob.getBytes());
    var pdfData = 'data:application/pdf;base64,' + base64data;

    // Generate a HTML output to automatically download the PDF
    var htmlOutput = HtmlService.createHtmlOutput(
      "<html><body>" +
      "<a id='pdfLink' href='" + pdfData + "' download='" + ss.getName() + ".pdf'></a>" +
      "<script>document.getElementById('pdfLink').click();</script>" +
      "</body></html>"
    );
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'File Downloaded Shortly');
  } else {
    Logger.log("Failed to fetch PDF with response code: " + response.getResponseCode());
  }

  // Clean up: remove created sheets and temporary spreadsheet
  createdSheets.forEach(function(sheet) {
    ss.deleteSheet(sheet);
  });
  DriveApp.getFileById(tempSsId).setTrashed(true);
}
