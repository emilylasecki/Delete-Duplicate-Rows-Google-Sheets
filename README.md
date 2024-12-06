# Delete-Duplicate-Rows-Google-Sheets
This Google Apps Script deletes duplicate rows from the last 4 days based on the day of execution. Whether a row is a duplicate or not is contingent on the content in cell D, that is student email. Note the script will incorrectly delete duplicates if the dates in column A are not sorted from oldest first to newest last.

## To Run
Navigate to the following link. Copy the content from the google sheet onto a seperate sheet on your own account. Make sure to delete blank rows from the bottom of the sheet so that only rows with data are present. Create a new Google Apps Script from the Extensions tab and copy the content of deleteDuplicates.js into a new script. Modify the sheet ID in lines 7 and 83 to match that of your Google Sheet. Configure with your Google account as prompted. Select run at the top of the Google Script to run the program.

https://docs.google.com/spreadsheets/d/1Bxpd4neH4HqJ1bRKU3Hp7sq7gc3JBZ12xS7USoFogUc/edit?gid=0#gid=0


