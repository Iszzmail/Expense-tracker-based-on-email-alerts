function syncHDFC_Monthly_Split() {
  // --- CONFIGURATION ---
  var searchQuery = 'from:alerts@hdfcbank.net "debited from" newer_than:5d';
  var cutoffDate = new Date(2025, 10, 20); // Nov 20, 2025
  // ---------------------

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var threads = GmailApp.search(searchQuery);

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    
    for (var j = 0; j < messages.length; j++) {
      var msg = messages[j];
      var body = msg.getPlainBody();

      // 1. DETERMINE SOURCE
      var source = "";
      // Check for Credit Card (Matches "Credit Card ending 7652" OR "Credit Card XX7652")
      if (body.match(/Credit Card (?:ending\s+)?7652/i) || body.match(/Credit Card\s+XX7652/i)) {
          source = "Credit Card";
      } 
      // Check for Bank Account (Matches "account 7616")
      else if (body.match(/account\s+7616/i)) {
          source = "Bank Account";
      } else {
          continue; // Skip emails that aren't for these two accounts
      }

      // 2. EXTRACT AMOUNT (Strict Logic)
      // We look for: "Rs." -> Number -> "has been debited" OR "is debited"
      var amountMatch = body.match(/Rs\.?\s*([0-9,]+(?:\.[0-9]{1,2})?)\s+(?:has\s+been|is)\s+debited/i);
      
      var amount = "0.00";
      if (amountMatch) {
          amount = amountMatch[1].replace(/,/g, ''); 
      } else {
          // If the specific phrase "Rs. X debited" isn't found, SKIP the email.
          // This ensures we never accidentally log "Rs. 8" from the footer.
          continue; 
      }

      // 3. EXTRACT DATE & TIME
      var dateString = "";
      var timeString = ""; 
      var txnDate = new Date(); 

      // CHECK: Is it the "Standard Alert" format? (Has "at HH:mm:ss")
      // Example: "on 22 Nov, 2025 at 22:30:42"
      var dateTimeMatch = body.match(/on\s+(\d{1,2}\s+\w{3},\s+\d{4})\s+at\s+(\d{2}:\d{2}:\d{2})/);
      
      if (dateTimeMatch) {
          dateString = dateTimeMatch[1]; 
          var rawTime = dateTimeMatch[2]; 
          timeString = rawTime.replace(/:/g, ""); 
          txnDate = new Date(dateString + " " + rawTime);
      } else {
          // Fallback: Try UPI format (e.g., "on 23-11-25")
          var dateMatch_upi = body.match(/on\s+(\d{2}-\d{2}-\d{2})/);
          if (dateMatch_upi) {
              dateString = dateMatch_upi[1];
              var parts = dateString.split("-");
              txnDate = new Date("20" + parts[2], parts[1] - 1, parts[0]);
              timeString = "000000"; // UPI emails usually don't have time in body
          } else {
              continue; // Skip if no date found
          }
      }

      if (txnDate < cutoffDate) {
        continue;
      }

      // 4. GENERATE ID (Ref No) & PARSE MERCHANT
      var refNo = "";
      var cleanMerchant = "Unknown";
      var fullDescription = "";

      var refMatch = body.match(/reference number is\s+(\d+)/);

      if (refMatch) {
          // --- CASE A: HAS REFERENCE NUMBER (UPI) ---
          refNo = refMatch[1]; 
          
          var merchantMatch = body.match(/to\s+(?:VPA\s+)?(.*?)\s+on/);
          fullDescription = merchantMatch ? merchantMatch[1].trim() : "Unknown";
          
          // Clean Name: Remove VPA/Email parts for Column C
          cleanMerchant = fullDescription.split(" ").filter(function(word) {
            return !word.includes("@") && !word.includes("."); 
          }).join(" ");

      } else {
          // --- CASE B: NO REFERENCE NUMBER (Standard) ---
          // Generate ID using Date + Time + Amount
          var dateCode = Utilities.formatDate(txnDate, Session.getScriptTimeZone(), "yyyyMMdd");
          refNo = "NO_REF_" + dateCode + "_" + timeString + "_" + amount;

          // Regex covers "towards X", "at X", or "to X"
          var merchantMatch = body.match(/(?:towards|at|to)\s+(.*?)\s+on/i);
          fullDescription = merchantMatch ? merchantMatch[1].trim() : "Unknown";
          cleanMerchant = fullDescription;
      }

      // --- WRITING TO SHEET ---
      var sheetName = Utilities.formatDate(txnDate, Session.getScriptTimeZone(), "MMM yyyy");
      var sheet = getOrCreateSheet(spreadsheet, sheetName);

      if (isDuplicate(sheet, refNo)) {
        continue;
      }
      
      // Formatting: Add apostrophe so Sheets treats RefNo as text
      var formattedRefNo = "'" + refNo;

      var dataRow = [dateString, source, cleanMerchant, amount, formattedRefNo, fullDescription];
      sheet.appendRow(dataRow);
      Logger.log("Added: " + cleanMerchant + " | Amt: " + amount);
    }
  }
}

// --- HELPER FUNCTIONS ---

function getOrCreateSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    var headers = [["Date", "Source", "Merchant", "Amount", "Reference No", "Full Description"]]; 
    sheet.getRange("A1:F1").setValues(headers).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  // Safety Check: If headers are missing, add them
  else if (sheet.getLastRow() === 0 || sheet.getRange("A1").getValue() !== "Date") {
    var headers = [["Date", "Source", "Merchant", "Amount", "Reference No", "Full Description"]]; 
    sheet.getRange("A1:F1").setValues(headers).setFontWeight("bold");
  }
  return sheet;
}

function isDuplicate(sheet, refNo) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false; 
  
  // Column E is index 5
  var data = sheet.getRange(2, 5, lastRow - 1, 1).getValues(); 
  
  // Clean Check: Remove apostrophe from existing data before comparing
  var existingRefs = data.map(function(r) { 
    return r[0].toString().replace(/^'/, ''); 
  });
  
  return existingRefs.indexOf(refNo) > -1;
}