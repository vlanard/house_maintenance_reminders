# Home Maintenance Reminder Script

A simple system that sends you email reminders for upcoming home maintenance tasks, managed in a Google Sheet.

## Quick Setup

### 1. Create Your Google Sheet
1. Make a copy of the Google Sheet [Template](https://docs.google.com/spreadsheets/d/1CWVOvbmrKbkni3k78G8RhObeJL5Ly3q1kw-DJyv3hvE/copy?usp=sharing) _including_ the script.

   NOTE: No data is shared. This script runs only on your sheet from your account.    
2. Populate the "house_log" sheet with your maintenance tasks and the frequency that you want a reminder each item is due.

### 2. Set Up the Script
1. If you made a copy of the sheet with the script, Go to **Extensions > Apps Script** from the menu bar. The script should be visible.
2. If you prefer to make your own Google sheet, Go to **Extensions > Apps Script**, then replace the default code completely with the contents of [Code.js](Code.js).
3. **Set your email** in `USER_CONFIG` at the top:
   ```javascript
   const USER_CONFIG = {
     email: "your.email@gmail.com"  // PUT YOUR EMAIL HERE
   };
   ```
4. Review the settings in `DEFAULT_CONFIG`, such as `triggerHours`, `triggerDays`, and `sheetName` in the `DEFAULT_CONFIG` to adjust or confirm.
5. Click **Save** (Ctrl+S)

### 3. Test & Schedule
1. Select `main` from the function dropdown next to the Debug button in the Apps Script Editor.
2. Click **Run** ▶️
3. Grant permissions when prompted. 
   - Note - This authorization will only grant YOU access to run YOUR own script against YOUR own sheet. No one else gets access.
   - When it says "Google hasn't verified this app", click Advanced
   - Then click `Go to (TEMPLATE) house maintenance (unsafe)` or similar.
   - check Select All, then click Continue. 
4. The script will run and you will see the output logs in your browser. You should get a test email with 1 sample item, plus any other items due.
   Once you verify the email was received, update the `Last Done` column for the task from the reminder with today's date.
5. Run `main()` once more to see how the completed item is no longer flagged as due.
6. Select `setupSchedule` from the function dropdown and click **Run**. This will set up the script on its automatic schedule. You're done!

#### Schedule changes
- Run `setupSchedule` once, whenever you want to change the schedule settings. 
- To turn off emails, run `endSchedule` once from the Apps Script Editor to delete the recurring job.

## Configuration Options

All configuration is done in `Code.js`. You can customize any setting in the `DEFAULT_CONFIG` object:

### Basic Settings
- `email`: Your email address (required - set in `USER_CONFIG`)
- `triggerDays`: Days to run (e.g., `["SATURDAY", "SUNDAY"]`) (default: `["SATURDAY"]`)
- `triggerHours`: Hours to run (e.g., `[7, 19]` for 7am and 7pm) (default: `[7]`)

### Sheet Settings  
- `sheetName`: Name of maintenance data sheet (default: `"house_log"`)
- `sheetRange`: Cell range to read from sheet (default: `"A1:L100"`)
- `columns`: Column header names that the script looks for

## Troubleshooting

- **No emails?** Check your email in `USER_CONFIG` and run `main()` manually
- **Permission errors?** Re-run the script and grant all requested permissions
- **Wrong schedule?** Run `setupSchedule()` again to update triggers
- **Missing columns?** Ensure your sheet has columns matching the expected names in `DEFAULT_CONFIG.columns`

That's it! Your home maintenance reminders are now automated. 