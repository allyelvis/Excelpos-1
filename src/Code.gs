function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() === 'Sales') {
    var editedCell = e.range;
    if (editedCell.getColumn() === 3) { // Quantity Sold column
      updateStock();
    }
  }
}

function updateStock() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sales');
  var range = sheet.getDataRange();
  var values = range.getValues();

  var productSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Products');
  var productRange = productSheet.getDataRange();
  var productValues = productRange.getValues();

  for (var i = 1; i < values.length; i++) {
    var productId = values[i][1];
    var quantitySold = values[i][2];

    for (var j = 1; j < productValues.length; j++) {
      if (productValues[j][0] == productId) {
        productValues[j][3] -= quantitySold; // Update stock
        break;
      }
    }
  }

  productRange.setValues(productValues);
}
