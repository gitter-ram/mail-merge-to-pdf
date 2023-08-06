# mail-merge-to-pdf
This is a small python script to split a PDF printed by word using mail merge and then send the split files to their intended recipients. Useful if you want to send certificates and the like to others in bulk.

# Installation and usage
1. Install python 3 and poppler. Add them to your PATH variable.
2. download the scipt present in this repository.
3. On windows we do not expect any errors to arise, for linux and Mac, install any modules that are not found by the sctipt.
4. Create an excel/calc spreadsheet with the mail merge data, and save it in the CSV format. Make sure not to use any commans in the records.
5. Copy the csv file to the same folder as this script.
6. Create the word document and do the mail merge as needed.
7. Print the mail merged document as PDF and save it in the same folder as this script.
8. Open the mailer.py script in notepad and enter the username and password as mentioned in the file.
9. Save the script.
10. run the script and answer the questions that follow.

# Troubleshooting
1. **"I have a G-Mail but there is a login error: application specific password ... What do I do?"**<br>
Go to https://myaccount.google.com/apppasswords and create an app password for your account. Then enter this password in the password section of the script.

2. **pdfseparate not found?**
Install Poppler from https://sourceforge.net/projects/poppler-win32/
