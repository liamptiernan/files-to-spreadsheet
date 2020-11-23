# files-to-spreadsheet
Google Apps Script to take CSV from drive and import to Google Sheets

Run the updateTabs() function to run the script.

Script searches through a specificed Google Drive folder ('driveFolder' variable) that contains .csv files. It creates a new sheet in the spreadsheet for each .csv and copies the CSV data into the new tab. 

Finally, it creates a query formula in the 'mainSheet' of the spread sheet to compile all of the CSV data together. This formula will exclude any sheet in 'protectedSheets' allowing you to prevent certain sheets from being compiled.

It will check if a file has already been brought into the sheet by checking the name against the existing tabs. If the file has already been processed, it will skip over it. As such, this script can be scheudled to run and will repeatedly check a folder for new files that have been uploaded.

Some specific ranges and cells in the code may need to be changed to suit the data you're working with. For instance, the content of the query formula will likely need to be updated to handle your specific data properly.

As well, the 'getWord()' function will likely need to be removed. This was built to get a specific word from a particular filename syntax. 