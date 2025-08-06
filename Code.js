/**
 * @fileoverview Automates the process of sending a daily report for the UPS Nogales account.
 * @version 1.1
 * @author Amado Trucking IT Department
 * @lastUpdated 2025-07-07
 *
 * @description
 * This script performs the following actions:
 * 1. Reads data from a source Google Sheet which contains a QUERY formula to filter records for 'UPS NOGALES'.
 * 2. Copies this filtered data into a new, temporary Google Sheet.
 * 3. Exports the temporary sheet as a Microsoft Excel file (.xlsx).
 * 4. Attaches the Excel file to an email and sends it to a predefined recipient.
 * 5. Deletes the temporary Google Sheet from Google Drive to avoid clutter.
 *
 * For complete process documentation, including the source formula and workflow, please refer to:
 * "IT Documentation: UPS Nogales Automated Report" 
 *  https://docs.google.com/document/d/16hT9A1UQLACYSDjdwj_0zhT2ajW8_nsH5RWZtbth8Q8/edit?tab=t.0
 */


function sendEmailWithExcelAttachment() {
  // Copy data to a new spreadsheet
  var newSpreadsheetId = copyDataToNewSpreadsheet();
  
  // Ensure all changes are applied
  SpreadsheetApp.flush();
  
  // Export the new spreadsheet as an Excel file
  var excelBlob = exportSpreadsheetAsExcel(newSpreadsheetId);
  
  // Send the Excel file as an email attachment
  var recipient = 'ups.chamberlain@amadotrucking.com'; // recipient's email address
  var subject = 'Amado - UPS Chamberlain Report';
  var body = 'Please find the attached Excel file containing the data.';
  
  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    body: body,
    attachments: [excelBlob]
  });
  
  // Delete the new spreadsheet to clean up
  DriveApp.getFileById(newSpreadsheetId).setTrashed(true);
}
/**************************************************************/
function copyDataToNewSpreadsheet() {
  var sourceSpreadsheetId = '1CM1xLzbtje6Lohy2KypvzkLFCouYNMZvIRxIz20QBqM'; // Google Sheets ID
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var sourceSheet = sourceSpreadsheet.getActiveSheet();
  var data = sourceSheet.getDataRange().getValues();
  
  // Create a new spreadsheet
  var newSpreadsheet = SpreadsheetApp.create('Copied Data');
  var newSheet = newSpreadsheet.getActiveSheet();
  
  // Paste data into the new spreadsheet
  newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  
  return newSpreadsheet.getId();
}


/**************************************************************/
function exportSpreadsheetAsExcel(spreadsheetId) {
  var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?format=xlsx';
  var params = {
    method: 'GET',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    }
  };
  var response = UrlFetchApp.fetch(url, params);

  // Get the current date
  var currentDate = new Date();
  
  // Format the date as "dd-MM-yyyy"
  var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'dd-MM-yyyy');
  
  // Set the blob's name with the formatted date
  var blob = response.getBlob().setName('UPS Chamberlain Report_' + formattedDate + '.xlsx');
  return blob;
}
