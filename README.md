Winlink Checkins Application - A Console Program to Tabulate Winlink Net Checkins
Description:
This application installs itself by default in C:\RMS Express\Global Folders\Messages but can be installed in any folder of your choosing. The net date range runs from midnight of the start date through midnight of the end date which are input when Checkins runs. It also asks for your callsign which is optional. If you provide the callsign, it will process message files in your personal folder ( C:\RMS Express\callsign\Messages). If you do not, it will run against messages in the folder where it was installed. You can get a copy at: https://drive.google.com/file/d/1lbN0yqKHRatzWRAnOHFdPWPKpwRZO2Wn/view?usp=sharing
Functionality:
The program will ask for a start date, an end date, and your callsign. It then looks in the specified/default folder for a file called “roster.txt” which is installed with the application. The first line in that file contains a unique keyword to identify the net (GLAWN) to be found in either the TO: field or the Message/Comment: field to work. It will scan the folder for MIME file types (the Winlink Messages) and process all files that fall between the start and end dates, looking for key words to identify the net, the callsign of the sender, and to identify the template, if used, in order to locate the checkin information. As it runs, some information is displayed in the console window to show status and events. The second line contains the roster list from the cell in red at the top of the spreadsheet. If there are new checkins, the roster.txt file will be updated, but the spreadsheet will require a manual update to add a new row for the new checkin callsign and information. 
Example:

When it is finished, it will deposit two files in the folder with the messages. The first is “checkins.txt” which contains a summary of all the records processed. The information in this file is used to populate a spreadsheet to keep track of the weekly checkins and administrative information. The second is “checkins.csv”, a comma delimited file with each row containing the elements of the checkin information. This can be reviewed to see the checkin basics and anything that was added as a comment.
Recommended Process:
Clear the Global Messages folder at the start of a new period (which could be a problem if you use it for something else). Clearing them away isn't necessary, but it can be useful to know how many physical files are going to be processed. One obvious advantage is if someone’s checkin is sent a few minutes early and is discarded by the application, it is easy to adjust the start date to catch it and not get something from the previous week. Save a copy of all of them in a folder outside of the Winlink Express structure. Everything should still work fine if the old messages are not removed.
For the coming week, move all of the new GLAWN messages to the Global Messages folder.
Run the Checkins program from the Start Menu or shortcut
Put the relevant data in the spreadsheet
Send the two files (checkins.txt and checkins.csv) to the GLAWN admins
Details:
Run the Checkins application
Enter the start and end dates (Both must be within 14 days of today):
Enter your callsign (or press enter for the default location):

These files are the output from the application:
checkins.txt
Checkins.csv
