# Expense Tracker: Gmail to Google Sheets

This project uses Google Apps Script to automatically scan your Gmail for HDFC transactions (RuPay Credit Card & Bank Account) and save them to a Google Sheet.

## Features
- **Centralized Tracking**: Writes all transactions to a single sheet named **"Expense tracker"**.
- **Multi-Account**: Tracks banks spending alerts emails and looks for keywords on the alert email For example, in the default script I used the
- "RuPay CC XX7652"
- "Account 7616"
- **Duplicate Prevention**: Uses the **UPI Transaction Reference Number** to ensure no duplicates.
- **Start Date**: Configured to start tracking from **Nov 20, 2025**.

## Setup Instructions

### 1. Prepare the Google Sheet
1. Create a new [Google Sheet](https://sheets.new).
2. The script will automatically create a tab named **"Expense tracker"** if it doesn't exist.

### 2. Add the Script
1. In your Google Sheet, go to **Extensions** > **Apps Script**.
2. Delete any code in the `Code.gs` file and paste the content of the `Code.gs` file from this project.
3. Save the project (Cmd/Ctrl + S).

### 3. Test the Script
1. Select the `processEmails` function and click **Run**.
2. Check your Google Sheet. You should see a tab named "Expense tracker" with your transactions!

### 4. Automate (Triggers)
1. In the Apps Script sidebar, click on **Triggers** (alarm clock icon).
2. Click **+ Add Trigger**.
3. Configure:
   - Function to run: `processEmails`
   - Event source: `Time-driven`
   - Type of time based trigger: `Hour timer`
   - Interval: `Every 2 hours`
4. Click **Save**.
