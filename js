var headerAppended = false;

// Function to append headers from Sheet1 to the Keepa sheet
function appendHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var keepaSheet = ss.getSheetByName("keepa");
  var sheet1 = ss.getSheetByName("Sheet1");
  var sheet1Headers = sheet1.getRange(1, 1, 1, sheet1.getLastColumn()).getValues()[0];
  var numHeaders = sheet1Headers.length;
  var remarkIndex = keepaSheet.getRange(1, 1, 1, keepaSheet.getLastColumn()).getValues()[0].indexOf("Remark");
  keepaSheet.insertColumnsBefore(remarkIndex + 1, numHeaders);
  keepaSheet.getRange(1, remarkIndex + 1, 1, numHeaders).setValues([sheet1Headers]);
  headerAppended = true;
}

// Function to process data and create ARRAYFORMULA formula
function processData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var keepaSheet = ss.getSheetByName("keepa");
  
  var keepaUPCEANColumn = Browser.inputBox("Enter the column letter for keepa UPC/EAN:");
  if (!keepaUPCEANColumn) {
    return;
  }

  var sheet1UPCColumn = Browser.inputBox("Enter the column letter for Sheet1 UPC:");
  if (!sheet1UPCColumn) {
    return;
  }

  var lastSheet1Column = ss.getSheetByName("Sheet1").getLastColumn();
  var lastSheet1Header = columnIndexToLetter(lastSheet1Column);

  var formula = "=ARRAYFORMULA(IF(IF(" + keepaUPCEANColumn + "2<>\"\", IFERROR(MATCH(" + keepaUPCEANColumn + "2, sheet1!" + sheet1UPCColumn + ":" + sheet1UPCColumn + ", 0), \"\"), \"\")<>\"\", INDEX(sheet1!A:" + lastSheet1Header + ", IF(" + keepaUPCEANColumn + "2<>\"\", IFERROR(MATCH(" + keepaUPCEANColumn + "2, sheet1!" + sheet1UPCColumn + ":" + sheet1UPCColumn + ", 0), \"\"), \"\")), \"\"))";
  
  keepaSheet.getRange("A2:A" + keepaSheet.getLastRow()).setFormula(formula);
  
  removeEmptyRows(keepaSheet);
}

// Function to convert column index to column letter
function columnIndexToLetter(columnIndex) {
  var dividend = columnIndex;
  var columnName = '';

  while (dividend > 0) {
    var modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = Math.floor((dividend - modulo) / 26);
  }

  return columnName;
}

// Function to remove rows with empty UPC values in Keepa sheet
function removeRowsWithEmptyUPC() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var keepaSheet = ss.getSheetByName("Keepa");
  
  var dataRange = keepaSheet.getDataRange();
  var values = dataRange.getValues();
  
  for (var i = values.length - 1; i >= 1; i--) {
    var upcEanValue = values[i][1];
    if (upcEanValue === "") {
      keepaSheet.deleteRow(i + 1);
    }
  }
}

function removeDuplicateRowsFromKeepa() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var keepaSheet = ss.getSheetByName("keepa"); // Change to your actual sheet name

  var dataRange = keepaSheet.getDataRange();
  var values = dataRange.getValues();

  var uniqueValues = [];
  var duplicateIndices = [];

  for (var i = 0; i < values.length; i++) {
    var value = values[i][1]; // Assuming the column index is 1 (column B)
    if (uniqueValues.indexOf(value) === -1) {
      uniqueValues.push(value);
    } else {
      duplicateIndices.push(i + 1); // Adding 1 to convert to 1-based index
    }
  }

  // Delete duplicate rows in reverse order to avoid index shifting
  for (var j = duplicateIndices.length - 1; j >= 0; j--) {
    keepaSheet.deleteRow(duplicateIndices[j]);
  }
}



function fillEmptyQtyWithOne() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Keepa"); // Get the "Keepa" sheet
  var dataRange = sheet.getDataRange(); // Get the range of all data in the sheet
  var values = dataRange.getValues(); // Get the values from the data range
  var qtyColumnIndex = values[0].indexOf("QTY"); // Find the column index of "QTY" header

  for (var i = 1; i < values.length; i++) { // Start from the second row (excluding header)
    var qtyValue = values[i][qtyColumnIndex]; // Get the value in the "QTY" column

    if (qtyValue === "" || isNaN(qtyValue)) {
      sheet.getRange(i + 1, qtyColumnIndex + 1).setValue(1);
    }
  }
}
// Function to remove duplicate rows based on UPC values in Keepa sheet
function applyColorBasedOnKeywordsAndQuantity() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Keepa"); // Get the "Keepa" sheet
  var dataRange = sheet.getDataRange(); // Get the range of all data in the sheet
  var values = dataRange.getValues(); // Get the values from the data range
  var titleColumnIndex = values[0].indexOf("Title"); // Find the column index of "Title" header
  var qtyColumnIndex = values[0].indexOf("QTY"); // Find the column index of "QTY" header

  if (titleColumnIndex !== -1 && qtyColumnIndex !== -1) { // Check if both columns exist
    // Define keyword patterns and number pattern
    var titleKeywords = ["pack of", "-pack", " pack", "pack", "/pk", "pk", "-pk"];
    var numberPattern = "\\b(\\d+)\\b";
    var regexPatterns = [];

    // Generate regex patterns for title keywords
    for (var i = 0; i < titleKeywords.length; i++) {
      regexPatterns.push(numberPattern + "\\s*" + titleKeywords[i], titleKeywords[i] + "\\s*" + numberPattern);
    }

    for (var j = 1; j < values.length; j++) { // Start from the second row (excluding header)
      var titleValue = values[j][titleColumnIndex].toString().toLowerCase(); // Get the value in the "Title" column
      var qtyValue = values[j][qtyColumnIndex]; // Get the value in the "QTY" column

      var titleNumber = null; // Initialize title number

      // Loop through regex patterns to find title number
      for (var k = 0; k < regexPatterns.length; k++) {
        var regex = new RegExp(regexPatterns[k], "i");
        var match = titleValue.match(regex);

        if (match) {
          titleNumber = parseFloat(match[1]);
          break;
        }
      }

      // If QTY is empty or not a number, fill with 1
      if (qtyValue === "" || isNaN(qtyValue)) {
        qtyValue = 1;
        sheet.getRange(j + 1, qtyColumnIndex + 1).setValue(qtyValue);
      }

      if (!isNaN(titleNumber)) { // Check if a title number was found
        // If title number and QTY are different, replace QTY with title number
        if (titleNumber !== qtyValue) {
          sheet.getRange(j + 1, qtyColumnIndex + 1).setValue(titleNumber);
        }

        var matchFound = false;

        // Loop through regex patterns to find a match
        for (var k = 0; k < regexPatterns.length; k++) {
          var regex = new RegExp(regexPatterns[k], "i");
          var match = titleValue.match(regex);

          if (match) {
            var keywordQty = parseFloat(match[1]);

            if (!isNaN(keywordQty) && keywordQty === qtyValue) {
              matchFound = true;
              break;
            }
          }
        }

        // Change cell color based on whether a match was found
        if (matchFound) {
          sheet.getRange(j + 1, titleColumnIndex + 1).setBackground("blue");
        } else {
          sheet.getRange(j + 1, titleColumnIndex + 1).setBackground("yellow");
        }
      }
    }
  }
}

function clearColors() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Keepa"); // Get the "Keepa" sheet
  var dataRange = sheet.getDataRange(); // Get the range of all data in the sheet
  dataRange.setBackground(null); // Clear background colors
}




function onEdit() {
  applyColorBasedOnKeywordsAndQuantity();
  fillEmptyQtyWithOne();
}

// Function to convert weight from kg to lb
function convertKgToLb(kg) {
  return kg /16; // Conversion factor
}

// Function to convert the weights in a specified column from kg to lb
function convertWeightsToLb(columnName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Keepa"); // Change to your sheet name
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var headerRow = values[0];
  
  var columnIndex = headerRow.indexOf(columnName);
  
  if (columnIndex === -1) {
    SpreadsheetApp.getUi().alert("Column '" + columnName + "' not found in header.");
    return;
  }

  // Loop through rows and convert weights
  for (var i = 1; i < values.length; i++) { // Start from the second row (excluding header)
    var weightKg = values[i][columnIndex];
    var weightLb = convertKgToLb(weightKg);
    values[i][columnIndex] = weightLb;
  }

  // Update the values in the sheet
  dataRange.setValues(values);
}

// You can call this function with the column name where weight is stored
convertWeightsToLb("Weight");

// Rest of your functions
function colorEmptyReferralAndFBAFees() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Keepa"); // Change to your sheet name
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var headerRow = values[0];

  var referralFeeIndex = headerRow.indexOf("Referral Fee %");
  var fbaFeeIndex = headerRow.indexOf("FBA Fee");

  if (referralFeeIndex === -1 || fbaFeeIndex === -1) {
    SpreadsheetApp.getUi().alert("Columns 'Referral Fee' and/or 'FBA Fee' not found in header.");
    return;
  }

  // Loop through rows and check for empty cells
  for (var i = 1; i < values.length; i++) { // Start from the second row (excluding header)
    if (values[i][referralFeeIndex] === "") {
      sheet.getRange(i + 1, referralFeeIndex + 1).setBackground("yellow");
    }
    if (values[i][fbaFeeIndex] === "") {
      sheet.getRange(i + 1, fbaFeeIndex + 1).setBackground("yellow");
    }
  }
}

// Call the function to apply the coloring
colorEmptyReferralAndFBAFees();

// Rest of your functions


//CALCULATE COG
function calculateCOGAndTotalCOG() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Keepa"); // Change to your sheet name
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var headers = values[0];
  
  var qtyColumnIndex = headers.indexOf("QTY");
  var productPriceColumnIndex = headers.indexOf("Price");
  var cogColumnIndex = headers.indexOf("COG"); // You can change this column name as needed
  var shippingColumnIndex = headers.indexOf("Shipping");
  var sendToAmazonColumnIndex = headers.indexOf("Send to Amazon");
  var prepColumnIndex = headers.indexOf("Prep");
  var totalCOGColumnIndex = headers.indexOf("Total COGS"); // Total COG column index
  
  if (
    qtyColumnIndex === -1 || 
    productPriceColumnIndex === -1 || 
    cogColumnIndex === -1 || 
    shippingColumnIndex === -1 || 
    sendToAmazonColumnIndex === -1 || 
    prepColumnIndex === -1 || 
    totalCOGColumnIndex === -1
  ) {
    SpreadsheetApp.getUi().alert("Required columns not found in header.");
    return;
  }
  
  for (var rowIndex = 1; rowIndex < values.length; rowIndex++) {
    var qtyValue = values[rowIndex][qtyColumnIndex];
    var productPriceValue = values[rowIndex][productPriceColumnIndex];
    var cogValue = isNaN(qtyValue) || isNaN(productPriceValue) ? 0 : qtyValue * productPriceValue;
    var shippingValue = values[rowIndex][shippingColumnIndex] || 0;
    var sendToAmazonValue = values[rowIndex][sendToAmazonColumnIndex] || 0;
    var prepValue = values[rowIndex][prepColumnIndex] || 0;
    
    var totalCOGValue = cogValue + shippingValue + sendToAmazonValue + prepValue;
    
    sheet.getRange(rowIndex + 1, cogColumnIndex + 1).setValue(cogValue);
    sheet.getRange(rowIndex + 1, totalCOGColumnIndex + 1).setValue(totalCOGValue);
  }
}

// Call the function to calculate COG and Total COG
calculateCOGAndTotalCOG();


function calculateMarginAndReferralFees() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Keepa"); // Change to your sheet name
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var headers = values[0];
  
  var amazonPriceColumnIndex = headers.indexOf("Amazon Price");
  var totalCogColumnIndex = headers.indexOf("Total COGS");
  var fbaFeeColumnIndex = headers.indexOf("FBA Fee");
  var marginColumnIndex = headers.indexOf("Net Margin");
  
  if (
    amazonPriceColumnIndex === -1 ||
    totalCogColumnIndex === -1 || 
    fbaFeeColumnIndex === -1 || 
    marginColumnIndex === -1
  ) {
    SpreadsheetApp.getUi().alert("Required columns not found in header.");
    return;
  }
  
  for (var rowIndex = 0; rowIndex < values.length; rowIndex++) {
    var amazonPriceValue = values[rowIndex][amazonPriceColumnIndex];
    var totalCogValue = values[rowIndex][totalCogColumnIndex];
    var fbaFeeValue = values[rowIndex][fbaFeeColumnIndex];
    
    if (
      typeof amazonPriceValue !== 'number' || 
      typeof totalCogValue !== 'number' || 
      typeof fbaFeeValue !== 'number'
    ) {
      continue; // Skip rows with non-numeric values
    }
    
    var marginValue = amazonPriceValue - totalCogValue - fbaFeeValue;
    
    // Set Margin value, replacing empty cells with 0
    sheet.getRange(rowIndex + 1, marginColumnIndex + 1).setValue(marginValue || 0);
  }
}

// Call the function to calculate Margin and Referral Fees
calculateMarginAndReferralFees();


function calculateMarginPercentageAndROI() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Keepa"); // Change to your sheet name
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var headers = values[0];
  
  var marginColumnIndex = headers.indexOf("Net Margin");
  var amazonPriceColumnIndex = headers.indexOf("Amazon Price");
  var marginPercentageColumnIndex = headers.indexOf("Net Margin%");
  var totalCOGColumnIndex = headers.indexOf("Total COGS");
  var roiColumnIndex = headers.indexOf("Net ROI%");
  
  if (
    marginColumnIndex === -1 || 
    amazonPriceColumnIndex === -1 || 
    marginPercentageColumnIndex === -1 || 
    totalCOGColumnIndex === -1 || 
    roiColumnIndex === -1
  ) {
    SpreadsheetApp.getUi().alert("Required columns not found in header.");
    return;
  }
  
  for (var rowIndex = 0; rowIndex < values.length; rowIndex++) {
    var marginValue = values[rowIndex][marginColumnIndex];
    var amazonPriceValue = values[rowIndex][amazonPriceColumnIndex];
    var totalCOGValue = values[rowIndex][totalCOGColumnIndex];
    
    if (
      typeof marginValue !== 'number' || 
      typeof amazonPriceValue !== 'number' || 
      typeof totalCOGValue !== 'number' || 
      totalCOGValue === 0
    ) {
      continue; // Skip rows with non-numeric values or zero Total COG
    }
    
    var marginPercentageValue = (marginValue / amazonPriceValue) * 100;
    var roiValue = (marginValue / totalCOGValue) * 100;
    
    // Set Margin Percentage value, replacing empty cells with 0
    sheet.getRange(rowIndex + 1, marginPercentageColumnIndex + 1).setValue(marginPercentageValue || 0);
    
   sheet.getRange(rowIndex + 1, roiColumnIndex + 1).setValue(roiValue || 0);
  }
}

// Call the function to calculate Margin Percentage and ROI
calculateMarginPercentageAndROI()




// Function to create custom menus in Google Sheets
function createMenus() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Keepa Menu')
    .addItem('Append Headers', 'appendHeaders')
    .addItem('Process Data', 'processData')
    .addItem('Remove Rows with Empty UPC', 'removeRowsWithEmptyUPC')
    .addItem('Remove Duplicate Rows', 'removeDuplicateRowsFromKeepa')
    .addItem('onEdit', 'onEdit')
    .addItem('convertKgToLb', 'convertKgToLb')
    .addItem('calculateCOGAndTotalCOG', 'calculateCOGAndTotalCOG')
    .addItem('calculateMarginAndReferralFees', 'calculateMarginAndReferralFees')
    .addItem('calculateMarginPercentageAndROI', 'calculateMarginPercentageAndROI')

    
    .addToUi();
}

// Function that runs when the spreadsheet is opened
function onOpen() {
  createMenus();
}
