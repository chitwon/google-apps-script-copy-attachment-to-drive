# Google apps script to copy email attachments to drive and log to spreadsheet.

I modified a script that copies gmail attachments to drive to also log the transaction to a spreadsheet. I used it for students to email me homework. 

Students enter a formatted subject line, email the attachment, and the script automatically parses the subject line, copies the attachment to drive, and appends an ongoing spreadsheet with the transaction. 

original code without spreadsheet by Andreas Gohr found here: https://www.splitbrain.org/blog/2017-01/30-save_gmail_attachments_to_google_drive


# Make it work for you.
First in gmail you have to create a filter for the subject line. The filter searches emails for a matching subject line, stars it, and labels it. 

From there, attach this script to a google spreadsheet. In the script editior, enable Resources > Advanced Google services enable and add Drive, Gmail, and Sheets. 

Then run main or schedule it through a script trigger. 
