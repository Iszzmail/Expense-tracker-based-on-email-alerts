function syncHDFC_Monthly_Split() {
  // --- CONFIGURATION ---
  // We search for BOTH Credit Card and Bank Account emails now
  var searchQuery = 'subject:"You have done a UPI txn. Check details!" newer_than:5d';
  
  // STRICT START DATE: Nov 20, 2025
  // (Month is 0-indexed in JS, so 10 = November)
  var cutoffDate = new Date(2025, 10, 20); 
  // ---------------------

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var threads = GmailApp.search(searchQuery);

  // We need to process messages to find out which month they belong to
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    
    for (var j = 0; j < messages.length; j++) {
      var msg = messages[j];
      var body = msg.getPlainBody();

      // 1. DETERMINE SOURCE (CC vs Bank)
      var source = "";
      if (body.indexOf("XX7652") > -1) {
        source = "Credit Card";
      } else if (body.indexOf("7616") > -1) {
        source = "Bank Account";
      } else {
        // Skip if it's neither (e.g. a different account)
        continue; 
      }

      // 2. EXTRACT DATE & FILTER
      // Pattern: "on 23-11-25" or "on 19-11-25"
      var dateMatch = body.match(/on\s+(\d{2}-\d{2}-\d{2})/);
      if (!dateMatch) continue;

      var dateString = dateMatch[1]; // e.g., "23-11-25"
      var dateParts = dateString.split("-"); // [23, 11, 25]
      // Create Date Object: Year (2025), Month (11-1 = 10), Day (23)
      var txnDate = new Date("20" + dateParts[2], dateParts[1] - 1, dateParts[0]);

      // STRICT CHECK: Ignore if before Nov 20, 2025
      if (txnDate < cutoffDate) {
        continue;
      }

      // 3. DETERMINE SHEET NAME (e.g., "Nov 2025")
      var sheetName = Utilities.formatDate(txnDate, Session.getScriptTimeZone(), "MMM yyyy");
      
      // Get or Create the specific monthly sheet
      var sheet = getOrCreateSheet(spreadsheet, sheetName);

      // 4. CHECK DUPLICATES IN THIS SPECIFIC SHEET
      // Ref No is now in Column E (Index 5)
      var refMatch = body.match(/reference number is\s+(\d+)/);
      var refNo = refMatch ? refMatch[1] : "";
      if (isDuplicate(sheet, refNo)) {
        continue;
      }

      // 5. EXTRACT REMAINING DATA
      // Amount
      var amountMatch = body.match(/Rs\.?(\d+(\.\d{1,2})?)/);
      var amount = amountMatch ? amountMatch[1] : "0.00";

      // Merchant (Clean logic for both CC and Bank)
      // Bank often uses "to VPA ... on", CC uses "to ... on"
      var merchantMatch = body.match(/to\s+(?:VPA\s+)?(.*?)\s+on/);
      var rawMerchant = merchantMatch ? merchantMatch[1].trim() : "Unknown";
      
      // Advanced Cleaning: Remove email-like IDs and "VPA" word
      var cleanMerchant = rawMerchant.split(" ").filter(function(word) {
        return !word.includes("@") && !word.includes("."); // Filter out upi-ids
      }).join(" ");
      
      // Fallback if cleaning removed everything (rare edge case)
      if (cleanMerchant.length < 2) cleanMerchant = rawMerchant;

      // 6. WRITE TO SHEET
      sheet.appendRow([dateString, source, cleanMerchant, amount, refNo, rawMerchant]);
      Logger.log("Added to " + sheetName + ": " + cleanMerchant + " (" + amount + ")");
    }
  }
}

// --- HELPER FUNCTIONS ---

function getOrCreateSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // Create Headers for the new month
    var headers = [["Date", "Source", "Merchant", "Amount", "Reference No", "Full Description"]];
    sheet.getRange("A1:F1").setValues(headers).setFontWeight("bold");
    // Optional: Freeze top row
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function isDuplicate(sheet, refNo) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false; // Only headers exist
  
  // Get all Reference Numbers (Column E is index 5) from row 2 down
  // Logic: getRange(row, col, numRows)
  var data = sheet.getRange(2, 5, lastRow - 1).getValues(); 
  // Flatten 2D array to 1D
  var existingRefs = data.map(function(r) { return r[0].toString(); });
  
  return existingRefs.indexOf(refNo) > -1;
}