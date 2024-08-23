function generateDivisionPalletPDF() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheetNames = sheets.map(function (sheet) {
    return sheet.getName();
  });


  var ui = SpreadsheetApp.getUi();
  var htmlOutput = HtmlService.createHtmlOutput('<html><body>' +
    '<style>' +
    'body { font-family: Arial, sans-serif; }' +
    'form { text-align: center; }' +
    'select { padding: 10px; font-size: 16px; margin-bottom: 20px; }' +
    'input[type="button"] { padding: 10px 20px; font-size: 16px; margin: 5px; cursor: pointer; }' +
    'input[type="button"]:hover { background-color: #ddd; }' +
    '</style>' +
    '<form id="sheetForm">' +
    '<label for="sheet">Select a sheet:</label><br><br>' +
    '<select id="sheet" name="sheet">' +
    sheetNames.map(function (name) {
      return '<option value="' + name + '">' + name + '</option>';
    }).join('') +
    '</select><br><br>' +
    '<input type="button" value="Submit" onclick="google.script.run.withSuccessHandler(closeDialog).processForm(document.getElementById(\'sheet\').value);google.script.host.close();">' +
    '<input type="button" value="Cancel" onclick="google.script.host.close()">' +
    '</form>' +
    '<script>' +
    'function closeDialog() { google.script.host.close(); }' +
    '</script>' +
    '</body></html>')
    .setWidth(300)
    .setHeight(200);
  ui.showModalDialog(htmlOutput, 'Select Sheet');
}


function processForm(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(sheetName);

  if (!sourceSheet) {
    Logger.log("Sheet '" + sheetName + "' not found");
    SpreadsheetApp.getUi().alert("Sheet '" + sheetName + "' not found. Please ensure the sheet exists.");
    return;
  }

  var data = sourceSheet.getDataRange().getValues();
  var createdSheets = [];
  var productTotals = {};

  var startRow = -1;
  var startColumn = -1;
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] && data[i][j].toString().toUpperCase() === 'PALLET TABLE') {
        startRow = i + 1;
        startColumn = j;
        break;
      }
    }
    if (startRow !== -1) break;
  }

  if (startRow === -1) {
    Logger.log("'Pallet Begins' not found");
    SpreadsheetApp.getUi().alert("'Pallet Begins' not found. Please ensure the keyword is present.");
    return;
  }

  var divisionColumn = startColumn + 3;
  var palletNoColumn = startColumn + 1;
  var customerNameColumn = startColumn + 2;

  var totalPallets = 0;

  // Find the maximum pallet number
  for (var i = startRow + 1; i < data.length; i++) {
    var palletNo = parseInt(data[i][palletNoColumn]);
    if (!isNaN(palletNo)) {
      totalPallets = Math.max(totalPallets, palletNo);
    }
  }

  var currentDivision = '';
  var currentCustomer = '';
  var currentPalletNo = '';
  var lastPalletProcessed = false;

  for (var i = startRow + 1; i < data.length; i++) {
    var row = data[i];

    if (row[divisionColumn].toUpperCase() === 'TOTAL') {
      continue;
    }

    if (row[divisionColumn] && row[divisionColumn].toUpperCase() !== 'TOTAL') {
      currentDivision = row[divisionColumn];
      currentPalletNo = row[palletNoColumn];
      currentCustomer = row[customerNameColumn];
      var sheetName = currentDivision + " " + currentPalletNo + " OF " + totalPallets;

      sheetName = sheetName.replace(/[\/:*?"<>|]/g, '-');

      var sheet = ss.getSheetByName(sheetName);

      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        createdSheets.push(sheet);
        sheet.appendRow([currentDivision + " " + currentPalletNo + " OF " + totalPallets]);
        sheet.appendRow([currentCustomer]);
        sheet.appendRow(['Product name', 'Quantity']);

        sheet.getRange('A1:B1').merge().setFontWeight('bold').setFontSize(24).setHorizontalAlignment('center').setVerticalAlignment('middle');
        sheet.getRange('A2:B2').merge().setFontWeight('bold').setFontSize(18).setHorizontalAlignment('center').setVerticalAlignment('middle');
        sheet.getRange('A3:B3').setFontWeight('bold').setFontSize(16).setBackground('#f2f2f2').setBorder(true, true, true, true, true, true).setHorizontalAlignment('center').setVerticalAlignment('middle');

        sheet.setColumnWidths(1, 2, 300);
        sheet.setRowHeights(1, 3, 50);
        sheet.setRowHeight(2, 90);

        // Center the content
        sheet.getRange('A1:B').setHorizontalAlignment('center');
      }
    }

    var totalQuantity = 0;
    for (var j = startColumn + 3; j < data[startRow].length; j++) {
      var productName = data[startRow][j];
      var quantity = row[j];
      if (productName && productName.toUpperCase() !== 'TOTAL' && quantity && !isNaN(quantity) && quantity > 0) {
        var newRow = sheet.appendRow([productName, quantity]);
        totalQuantity += quantity;

        // Update the product totals
        if (productTotals[productName]) {
          productTotals[productName] += quantity;
        } else {
          productTotals[productName] = quantity;
        }

        var lastRow = sheet.getLastRow();
        sheet.setRowHeight(lastRow, 60);
        var range = sheet.getRange('A' + lastRow + ':B' + lastRow);
        range.setBorder(true, true, true, true, true, true).setFontWeight('normal').setBackground(null).setHorizontalAlignment('center').setVerticalAlignment('middle');
        if (lastRow % 2 === 0) {
          range.setBackground('#f9f9f9');
        } else {
          range.setBackground('#ffffff');
        }
        sheet.getRange('A' + lastRow).setFontSize(14);
        sheet.getRange('B' + lastRow).setFontSize(14);
        sheet.setRowHeight(lastRow, 40);
      }
    }

    var totalRow = sheet.appendRow(['TOTAL', totalQuantity]);
    var lastRow = sheet.getLastRow();
    var totalRange = sheet.getRange('A' + lastRow + ':B' + lastRow);
    totalRange.setFontWeight('bold').setFontSize(16).setBackground('#f2f2f2').setBorder(true, true, true, true, true, true).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(lastRow, 50);

    if (parseInt(currentPalletNo) == totalPallets) {
      lastPalletProcessed = true;
      break;
    }
  }

  var productOrder = [
    "Idli", "Dosa", "Idli Family", "Dosa Family", "Idli-Dosa Party 2.0",
    "Dosa Party 2.0", "Jumbo Idli", "Jumbo Dosa", "Organic Idli",
    "Organic Dosa", "Methi Dosa", "Millet Dosa", "Yellow Lentil Dosa",
    "Uthappam", "Adai", "Pesarattu", "Ragi Dosa", "Sambar",
    "Coconut Sambar", "Moong Dhal Sambar", "Mango Dhal", "Rasam",
    "Lemon Pickles", "Mango Pickles", "Tomato Thokku",
    "Biriyani Paste", "Chkn Curry Paste"
  ];

  // Create the Summary sheet
  var summarySheet = ss.insertSheet('Summary');
  summarySheet.appendRow(['Summary']);
  summarySheet.appendRow(['Product', 'Quantity']);

  summarySheet.getRange('A1:B1').merge().setFontWeight('bold').setFontSize(24)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  summarySheet.getRange('A2:B2').setFontWeight('bold').setFontSize(18).setBorder(true, true, true, true, true, true)
    .setBackground('#f2f2f2').setHorizontalAlignment('center').setVerticalAlignment('middle');

  summarySheet.setColumnWidths(1, 2, 300);
  summarySheet.setRowHeights(1, 3, 150);
  summarySheet.setRowHeight(2, 50);

  var totalQuantitySum = 0;

  // Iterate over the products in the custom order
  for (var i = 0; i < productOrder.length; i++) {
    var product = productOrder[i];
    var quantity = productTotals[product] || 0;

    // Only add to the summary sheet if the quantity is greater than 0
    if (quantity > 0) {
      summarySheet.appendRow([product, quantity]);

      // Update the total quantity sum
      totalQuantitySum += quantity;

      var lastRow = summarySheet.getLastRow();
      summarySheet.getRange('A' + lastRow + ':B' + lastRow).setBorder(true, true, true, true, true, true)
        .setFontWeight('normal').setFontSize(16).setBackground(null)
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
      summarySheet.setRowHeight(lastRow, 40);
    }
  }

  // Insert the Total Quantity row
  summarySheet.appendRow(['TOTAL', totalQuantitySum]);

  var lastRow = summarySheet.getLastRow();
  summarySheet.getRange('A' + lastRow + ':B' + lastRow).setBorder(true, true, true, true, true, true)
    .setFontWeight('bold').setFontSize(16).setBackground('#f2f2f2')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  summarySheet.setRowHeight(lastRow, 50);




  createdSheets.push(summarySheet);


  if (lastPalletProcessed) {
    var tempSpreadsheet = SpreadsheetApp.create('Temp Spreadsheet for PDF');
    var tempSsId = tempSpreadsheet.getId();
    var tempSs = SpreadsheetApp.openById(tempSsId);


    createdSheets.forEach(function (sheet) {
      sheet.copyTo(tempSs).setName(sheet.getName());
    });


    var tempSheet = tempSs.getSheetByName('Sheet1');
    if (tempSheet) {
      tempSs.deleteSheet(tempSheet);
    }


    var url = 'https://docs.google.com/spreadsheets/d/' + tempSsId + '/export?';
    var url_ext = 'exportFormat=pdf&format=pdf' +
      '&size=A4' +
      '&portrait=true' +
      '&fitw=true' +
      '&sheetnames=false&printtitle=false' +
      '&pagenumbers=false&gridlines=false' +
      '&fzr=false' +
      '&top_margin=0.75' +
      '&bottom_margin=0.75' +
      '&left_margin=1.5' +  // Increased left margin
      '&right_margin=1.5';  // Increased right margin


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


    createdSheets.forEach(function (sheet) {
      ss.deleteSheet(sheet);
    });
    DriveApp.getFileById(tempSsId).setTrashed(true);
  }
}







