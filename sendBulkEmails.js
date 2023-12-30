function sendBulkEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Your Sheet Name');
  const data = sheet.getRange('H2:H' + sheet.getLastRow()).getValues().flat();
  // change H2:H to your email column address

  // Subject and body of the email
  let subject = 'Hii, This is Raushan';
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
