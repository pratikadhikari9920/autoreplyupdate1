function extractFieldsFromEmail() {
  var sheet = SpreadsheetApp.openById("1MDkg-T-pyIuG1HvK6Oii7mZ35pv8IW40kgQKqfocrmA").getActiveSheet();
  var serialSheet = SpreadsheetApp.openById("1MDkg-T-pyIuG1HvK6Oii7mZ35pv8IW40kgQKqfocrmA").getSheetByName("Sheet2");

  // Check if the main sheet has headers; if not, add them
  if (sheet.getLastRow() === 0) {
    var headers = ['Name', 'Mobile No.', 'Address', 'District', 'Pincode', 'Country', 'Email used for purchasing books on online platforms', 'Serial Number', 'Sent To'];
    sheet.appendRow(headers); // Append the header row
  }

  // Load all serial numbers into an array from Sheet2
  var serialNumbers = serialSheet.getRange("A2:A").getValues().flat();
  var serialIndex = 0;

  // Get all existing data in the main sheet to check for duplicates
  var data = sheet.getDataRange().getValues();
  var existingMobileNumbers = new Set();
  var existingEmails = new Set();
  var usedSerialNumbers = new Set();

  for (var i = 1; i < data.length; i++) {
    existingMobileNumbers.add(data[i][1]); // Mobile No.
    existingEmails.add(data[i][8]); // "Sent To" column
    usedSerialNumbers.add(data[i][7]); // Serial Number column
  }

  var threads = GmailApp.getInboxThreads(0, 50);

  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    var messages = thread.getMessages();

    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var body = message.getPlainBody();
      var fromEmail = message.getFrom().match(/<(.+)>/)[1]; // Extract only the email part

      // Skip if the email has already received a serial number
      if (existingEmails.has(fromEmail)) {
        Logger.log('Duplicate entry found for email: ' + fromEmail);
        continue;
      }

      try {
        var name = extractField(body, 'Name:');
        var mobile = extractField(body, 'Mobile No.:');
        var address = extractField(body, 'Address:');
        var district = extractField(body, 'District:');
        var pincode = extractField(body, 'Pincode:');
        var country = extractField(body, 'Country:');
        var emailUsed = extractField(body, 'Email used for purchasing books on online platforms:');

        // Ensure all fields are filled out
        if (!name || !mobile || !address || !district || !pincode || !country || !emailUsed) {
          Logger.log('Required fields are missing in the email from: ' + fromEmail);
          continue;
        }

        // Check for duplicates based on mobile number or purchase email
        if (existingMobileNumbers.has(mobile) || existingEmails.has(emailUsed)) {
          Logger.log('Duplicate entry found for mobile number or email: ' + fromEmail);
          continue;
        }

        // Ensure there are still serial numbers left
        if (serialIndex < serialNumbers.length) {
          // Check if the serial number is already used
          while (serialIndex < serialNumbers.length && usedSerialNumbers.has(serialNumbers[serialIndex])) {
            serialIndex++; // Skip used serial numbers
          }

          if (serialIndex < serialNumbers.length) {
            var serialNumber = serialNumbers[serialIndex++];
            var row = [name, mobile, address, district, pincode, country, emailUsed, serialNumber, fromEmail];
            sheet.appendRow(row);

            GmailApp.sendEmail(emailUsed, "Your Serial Number", "Thank you for your purchase! Your serial number is: " + serialNumber);

            var emailColumnIndex = 7; // "Serial Number" column index
            var lastRow = sheet.getLastRow();
            sheet.getRange(lastRow, emailColumnIndex + 2).setBackground('green'); // Set "Sent To" column background to green

            // Add this email and serial number to the existing sets to prevent duplicates
            existingMobileNumbers.add(mobile);
            existingEmails.add(emailUsed);
            usedSerialNumbers.add(serialNumber);
          } else {
            Logger.log('No more available serial numbers to assign.');
          }
        } else {
          Logger.log('No more serial numbers available for email: ' + emailUsed);
        }

      } catch (error) {
        Logger.log('Error processing email from: ' + fromEmail + ' - ' + error.message);
      }
    }
  }
}

// Helper function to extract specific fields from the email body
function extractField(body, fieldName) {
  var regex = new RegExp(fieldName + '\\s*(.*)', 'i');
  var match = body.match(regex);
  return match ? match[1].trim() : null;
}

function createTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'extractFieldsFromEmail') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('extractFieldsFromEmail')
    .timeBased()
    .everyMinutes(1) // Set to run every 1 minute
    .create();
}
