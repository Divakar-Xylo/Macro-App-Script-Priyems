# Generate Styled PDF from Google Sheets

## Overview
This script automates the creation of new sheets based on data from a source sheet named "Sheet1" in a Google Spreadsheet. It then exports all the created sheets to a single styled PDF and provides an automatic download link for the PDF. Finally, it removes the created sheets after generating the PDF.

## Prerequisites
1. A Google Spreadsheet with a source sheet named "Sheet1".
2. Google Apps Script enabled for the Google Spreadsheet.

## Setup and Execution

### Step 1: Open Google Sheets
1. Open the Google Sheets document where you want to run this script.

### Step 2: Open Script Editor
1. Click on `Extensions` in the menu.
2. Select `Apps Script`.

### Step 3: Add the Script
1. Delete any existing code in the script editor.
2. Copy and paste the code from `macro.js` into the script editor.

### Step 4: Save the Script
1. Click on the floppy disk icon or press `Ctrl + S` to save the script.
2. Name the script, e.g., "GenerateStyledPDF".

### Step 5: Authorize the Script
1. Click on the `Run` button (the triangular "play" icon).
2. You will be prompted to authorize the script to access your Google Sheets and Google Drive.
3. Follow the authorization steps and grant the required permissions.

### Step 6: Run the Script
1. Once authorized, click the `Run` button again to execute the script.
2. The script will:
    - Check if "Sheet1" exists.
    - Create new sheets based on the divisions, pallet numbers, and customers from the source data.
    - Style the new sheets.
    - Export all sheets to a single PDF.
    - Provide an automatic download link for the PDF.
    - Delete the newly created sheets after generating the PDF.

### Step 7: Verify the PDF
1. After the script runs successfully, a dialog box will appear with a download link for the PDF.
2. Click the link to download and verify the PDF.

## Troubleshooting
- **Sheet "Sheet1" not found**: Ensure that the source sheet is named exactly "Sheet1".
- **Authorization issues**: Ensure you have granted all required permissions.
- **Script errors**: Check the log in the script editor for any error messages and debug accordingly.

## Notes
- Ensure your data in "Sheet1" is well-organized and follows the expected format.
- The script merges cells and styles headers for better readability.
- Created sheets are automatically deleted after the PDF generation to keep the spreadsheet clean.

By following these steps, you can successfully run the script to create styled PDFs from your Google Sheets data.
