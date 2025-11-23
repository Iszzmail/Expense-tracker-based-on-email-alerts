/**
 * Google Apps Script to extract RuPay transaction details from Gmail and append to Google Sheets.
 * 
 * SETUP INSTRUCTIONS:
 * 1. Create a Google Sheet.
 * 2. Add headers in Row 1: "Date", "Merchant", "Amount", "Debit/Credit", "Description", "Message ID".
 * 3. Go to Extensions > Apps Script.
 * 4. Paste this code.
 * 5. Run 'setupLabel' function once.
 * 6. Run 'processEmails' to test.
 * 7. Set up a Time-driven trigger for 'processEmails' to run every 2 hours.
 */

// CONFIGURATION
var SHEET_NAME = "Sheet1"; // Name of your sheet tab
var SEARCH_QUERY = 'subject:"RuPay" -label:ExpenseTrackerProcessed'; // Search query for emails
var PROCESSED_LABEL_NAME = "ExpenseTrackerProcessed";

function processEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log("Sheet not found: " + SHEET_NAME);
    return;
  }

  // Create label if it doesn't exist
  createLabelIfNeeded(PROCESSED_LABEL_NAME);
  var label = GmailApp.getUserLabelByName(PROCESSED_LABEL_NAME);

  // Search for emails
  var threads = GmailApp.search(SEARCH_QUERY);
  Logger.log("Found " + threads.length + " threads.");

  var newTransactions = [];

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      
      // Skip if already processed (double check, though search query handles this)
      // Note: Gmail search isn't always instant with labels, so good to be safe if running frequently
      
      var body = message.getPlainBody();
      var subject = message.getSubject();
      var date = message.getDate();
      var msgId = message.getId();

      var transaction = parseTransaction(body, subject, date, msgId);
      
      if (transaction) {
        newTransactions.push([
          transaction.date,
          transaction.merchant,
          transaction.amount,
          transaction.type,
          transaction.description,
          transaction.msgId
        ]);
      }
    }
    
    // Mark thread as processed
    threads[i].addLabel(label);
  }

  // Append to sheet
  if (newTransactions.length > 0) {
    // Sort by date if needed, or just append
    // newTransactions.sort(function(a, b) { return a[0] - b[0]; });
    
    sheet.getRange(sheet.getLastRow() + 1, 1, newTransactions.length, newTransactions[0].length).setValues(newTransactions);
    Logger.log("Added " + newTransactions.length + " transactions.");
  } else {
    Logger.log("No new transactions found.");
  }
}

function parseTransaction(body, subject, date, msgId) {
  // CUSTOMIZE THIS REGEX BASED ON YOUR BANK'S EMAIL FORMAT
  // Example pattern for many banks: "Rs. 123.00 spent on RuPay card XX1234 at MERCHANT NAME on 01-01-2025"
  
  var amount = 0;
  var merchant = "Unknown";
  var type = "Debit"; // Default to Debit
  
  // Regex examples - YOU MAY NEED TO ADJUST THESE
  // Look for "Rs." or "INR" followed by digits
  var amountRegex = /(?:Rs\.?|INR)\s*([\d,]+\.?\d*)/i;
  var merchantRegex = /(?:at|to)\s+([A-Za-z0-9\s\.\-\*]+?)(?:\s+on|\s+using|\.$)/i;
  
  var amountMatch = body.match(amountRegex);
  var merchantMatch = body.match(merchantRegex);

  if (amountMatch) {
    amount = amountMatch[1].replace(/,/g, ''); // Remove commas
  }
  
  if (merchantMatch) {
    merchant = merchantMatch[1].trim();
  }
  
  // If we couldn't find an amount, it might not be a transaction email
  if (!amountMatch) {
    return null;
  }

  return {
    date: date,
    merchant: merchant,
    amount: amount,
    type: type,
    description: subject,
    msgId: msgId
  };
}

function setupLabel() {
  createLabelIfNeeded(PROCESSED_LABEL_NAME);
  Logger.log("Label setup complete.");
}

function createLabelIfNeeded(name) {
  var label = GmailApp.getUserLabelByName(name);
  if (!label) {
    GmailApp.createLabel(name);
  }
}
