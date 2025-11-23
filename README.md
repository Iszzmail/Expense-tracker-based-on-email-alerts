# Expense Tracker: Gmail to Google Sheets

This project uses Google Apps Script to automatically scan your Gmail for RuPay card transactions and save them to a Google Sheet.

## Features
- **Automatic Sync**: Runs automatically every 2 hours (or your preferred interval).
- **Duplicate Prevention**: Uses a Gmail label (`ExpenseTrackerProcessed`) to ensure emails aren't counted twice.
- **Customizable**: You can adjust the parsing logic to match your specific bank's email format.

## Setup Instructions

### 1. Prepare the Google Sheet
1. Create a new [Google Sheet](https://sheets.new).
2. Rename the tab at the bottom to `Sheet1` (if it isn't already).
3. Add the following headers to the first row (A1:F1):
   - `Date`
   - `Merchant`
   - `Amount`
   - `Debit/Credit`
   - `Description`
   - `Message ID`

### 2. Add the Script
1. In your Google Sheet, go to **Extensions** > **Apps Script**.
2. Delete any code in the `Code.gs` file and paste the content of the `Code.gs` file from this project.
3. Save the project (Cmd/Ctrl + S).

### 3. Test the Script
1. In the Apps Script toolbar, select the `setupLabel` function and click **Run**.
   - You will need to authorize the script to access your Gmail and Sheets.
2. Send yourself a test email with the subject "RuPay" and body "Rs. 100.00 spent at TestMerchant".
3. Select the `processEmails` function and click **Run**.
4. Check your Google Sheet. You should see the transaction!
5. Check your Gmail. The email should now have the label `ExpenseTrackerProcessed`.

### 4. Automate (Triggers)
1. In the Apps Script sidebar, click on **Triggers** (alarm clock icon).
2. Click **+ Add Trigger**.
3. Configure:
   - Function to run: `processEmails`
   - Event source: `Time-driven`
   - Type of time based trigger: `Hour timer`
   - Interval: `Every 2 hours`
4. Click **Save**.

## Customization
If your bank emails look different, you may need to edit the `parseTransaction` function in `Code.gs`. Look for the `Regex examples` section and adjust the patterns to match your email body.
