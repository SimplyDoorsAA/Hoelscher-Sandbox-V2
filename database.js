function doGet(e) {
  try {
    // Connects to the first sheet in your workbook
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    // Security Authorization
    const MULTIPLIER = 0.48;
    const INTERNAL_PIN = "5588"; 
    const isInternalRequest = (e.parameter.pin === INTERNAL_PIN);
    
    const catalogData = rows.map(row => {
      const item = {};
      headers.forEach((header, index) => {
        item[header] = row[index];
        // Only calculate and attach NET costs if PIN is valid
        if (typeof row[index] === "number" && isInternalRequest) {
          item[`${header}_Net`] = parseFloat((row[index] * MULTIPLIER).toFixed(2));
        } else if (typeof row[index] === "number" && !isInternalRequest) {
          // If no PIN, map the List price to Net to hide margins
          item[`${header}_Net`] = parseFloat(row[index].toFixed(2));
        }
      });
      return item;
    });
    
    const payload = { 
      success: true, 
      data: catalogData, 
      internalMode: isInternalRequest,
      // Passes the multiplier securely to the frontend for Add-ons only if authorized
      multiplier: isInternalRequest ? MULTIPLIER : 1.0 
    };
    
    const output = ContentService.createTextOutput(JSON.stringify(payload));
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
    
  } catch (error) {
    const errorPayload = { error: error.toString(), success: false };
    return ContentService.createTextOutput(JSON.stringify(errorPayload)).setMimeType(ContentService.MimeType.JSON);
  }
}
