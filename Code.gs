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
      if (body.match(/Credit Card (?:ending\s+)?7652/i) || body.match(/Credit Card\s+XX7652/i)) {
          source = "Credit Card";
      } else if (body.match(/account\s+7616/i)) {
          source = "Bank Account";
      } else {
          continue; 
      }

      // 2. DETERMINE IF UPI
      var isUPI = (body.indexOf("Your UPI transaction reference number is") > -1);

      // 3. EXTRACT CORE DATA
      var amountMatch = body.match(/Rs\.?(\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?)/);
      var amount = amountMatch ? amountMatch[1].replace(/,/g, '') : "0.00"; 

      // Date Parsing
      var dateMatch_standard = body.match(/on\s+(\d{1,2}\s+\w{3},\s+\d{4})/);
      var dateMatch_upi = body.match(/on\s+(\d{2}-\d{2}-\d{2})/);
      var dateString = dateMatch_standard ? dateMatch_standard[1] : (dateMatch_upi ? dateMatch_upi[1] : "");

      var txnDate = parseDateString(dateString); 
      if (txnDate < cutoffDate) {
        continue;
      }
      
      var refNo = "";
      var cleanMerchant = "Unknown";
      var fullDescription = ""; 

      if (isUPI) {
          // --- UPI LOGIC ---
          var refMatch = body.match(/reference number is\s+(\d+)/);
          refNo = refMatch ? refMatch[1] : "";
          
          var merchantMatch = body.match(/to\s+(?:VPA\s+)?(.*?)\s+on/);
          fullDescription = merchantMatch ? merchantMatch[1].trim() : "Unknown";
          
          cleanMerchant = fullDescription.split(" ").filter(function(word) {
            return !word.includes("@") && !word.includes("."); 
          }).join(" ");

      } else {
          // --- STANDARD ALERT LOGIC (Fixed Regex) ---
          
          // Generate a stable ID for deduplication
          // We use Date + Source + Amount. This prevents "6 times" duplication.
          var uniqueDateCode = Utilities.formatDate(txnDate, Session.getScriptTimeZone(), "yyyyMMdd");
          refNo = "NO_REF_" + uniqueDateCode + "_" + source.replace(" ", "") + "_" + amount;
          
          // Improved Regex: Matches "towards X on", "at X on", or "to X on"
          var merchantMatch = body.match(/(?:towards|at|to)\s+(.*?)\s+on/i);
          fullDescription = merchantMatch ? merchantMatch[1].trim() : "Unknown";
          cleanMerchant = fullDescription; 
      }

      // --- WRITING DATA ---
      var sheetName = Utilities.formatDate(txnDate, Session.getScriptTimeZone(), "MMM yyyy");
      var sheet = getOrCreateSheet(spreadsheet, sheetName);

      // Check Duplicates (Now strips apostrophes for accurate checking)
      if (isDuplicate(sheet, refNo)) {
        continue;
      }
      
      // Add apostrophe for display, but isDuplicate handles this now
      var formattedRefNo = (refNo.indexOf("NO_REF") === -1) ? "'" + refNo : refNo;

      // COLUMNS: Date | Source | Merchant | Amount | Reference No | Full Description
      var dataRow = [dateString, source, cleanMerchant, amount, formattedRefNo, fullDescription];
      sheet.appendRow(dataRow);
      Logger.log("Added: " + cleanMerchant);
    }
  }
}

// --- HELPER FUNCTIONS ---

function parseDateString(dateStr) {
    if (!dateStr) return new Date(0);
    if (dateStr.match(/^\d{2}-\d{2}-\d{2}$/)) {
        var parts = dateStr.split("-");
        return new Date("20" + parts[2], parts[1] - 1, parts[0]);
    }
    if (dateStr.match(/^\d{1,2}\s+\w{3},\s+\d{4}$/)) {
        return new Date(dateStr);
    }
    return new Date(0); 
}

function getOrCreateSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // FORCE HEADERS on creation
    var headers = [["Date", "Source", "Merchant", "Amount", "Reference No", "Full Description"]]; 
    sheet.getRange("A1:F1").setValues(headers).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  // SAFETY CHECK: If sheet exists but is empty/missing headers, add them
  else if (sheet.getLastRow() === 0 || sheet.getRange("A1").getValue() !== "Date") {
    var headers = [["Date", "Source", "Merchant", "Amount", "Reference No", "Full Description"]]; 
    sheet.getRange("A1:F1").setValues(headers).setFontWeight("bold");
  }
  return sheet;
}

function isDuplicate(sheet, refNo) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false; 
  
  // Get existing refs from Column E
  var data = sheet.getRange(2, 5, lastRow - 1, 1).getValues(); 
  
  // CLEAN CHECK: Remove apostrophe before comparing
  var existingRefs = data.map(function(r) { 
    return r[0].toString().replace(/^'/, ''); 
  });
  
  return existingRefs.indexOf(refNo) > -1;
}