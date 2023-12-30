//This function creat a menu in the spreadsheet
function onOpen() {
  createMenuWithSubMenu();
}

function createMenuWithSubMenu() {
// creating some sub-menus  
  SpreadsheetApp.getUi().createMenu("ε(´｡•᎑•`)っ 💕")
    .addItem("🧪 Get Result", "filterAndCopyData")
    .addSeparator()
    .addItem("🐾 Clear", "clearData")
    .addSeparator()
    .addItem("📩 Sent Bulk Email", "sendBulkEmails")
    .addSeparator()
    .addToUi();
}
// bulk email sending function starts here
function sendBulkEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('your sheet name');
  const data = sheet.getRange('H2:H10' + sheet.getLastRow()).getValues().flat();
  // don't forget to change your email list column  to H2:H10

  // Subject and body of the email
  let subject = 'Hii, Your Webmail Disk Full: Clean it Immediately';
  let body = "Umm,\n\nI hope you are doing well.\n\nHey what are you doing?\n\ni'm sure you are doing nothing\nthanks for your attention";

  // Loop through each email address and send the email
  for (var i = 0; i < data.length; i++) {
    var emailAddress = data[i];
    
    // Check if the email address is not empty
    if (emailAddress && emailAddress.trim() !== "") {
      sendEmail(emailAddress, subject, body);
    } else {
      Logger.log('Skipping empty email address at row ' + (i + 2));
    }
  }
}

function sendEmail(emailAddress, subject, body) {
  try {
    GmailApp.sendEmail(emailAddress, subject, body);
    Logger.log('Email sent to: ' + emailAddress);
  } catch (error) {
    Logger.log('Error sending email to ' + emailAddress + ': ' + error.toString());
  }
}
//This function clear the spreadsheet with their formatting
function clearData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('your sheet name');
  var resultClear = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('your sheet name');
  
  // Clear content in specified ranges
  // you can change your clear range data
  sheet.getRange('A3:E').clearContent();
  resultClear.getRange('A2:C').clearContent();

  // Apply formatting changes
  var headerRange = sheet.getRange('A3:E');

  // Modify these style attributes based on your preference
  headerRange.setBackground('#ffffff').setFontColor('#000000').setHorizontalAlignment('left').setFontSize(10);
 
}

//This function works with regular formulas and then it works
//This function will used to copy and modify created data
function filterAndCopyData() {
  var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Calculation');
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Result');

  // Get data from B2:D
  var lastRow = sourceSheet.getLastRow();
  var dataRange = sourceSheet.getRange('B2:D' + lastRow);
  var dataValues = dataRange.getValues();

  // Sort data based on column D
  dataValues.sort(function (a, b) {
    var order = { 'GB': 1, 'MB': 2, 'KB': 3, 'Bytes': 4 };
    return order[a[2]] - order[b[2]];
  });

  // Clear existing content in the target sheet
 // targetSheet.clear();

  // Paste the sorted data starting from the second row in the "Result" sheet
  targetSheet.getRange(2, 1, dataValues.length, 3).setValues(dataValues);
}
