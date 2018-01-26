# SLA Excel Filename updater
## Saves you time with looking at tons of dates!
For an SLA job at work, we have to have paperwork set up with a specific naming convention with dates that change every week. Doing that by hand is tedious, so I made a program to do it for me.

An example of the file names is "Kersten, William - CSCI 1010 W1S Spring 2018 - Attendance 1-22 to 1-26". Everything in this filename will stay the same through one semester, except for the dates at the end. This program basically grabs those dates, adds 7 to them, and plaps 'em back in.

## How To Use
### NOTE: You must already have the Excel documents created and with the proper names! Check the example file name above to make sure you have correctly set the names up, otherwise this WILL NOT WORK.

1. Download this repository as a zip (green button above, to the right. Click it, then Download Zip.) 
2. Move that .zip file to wherever you are keeping your SLA Excel files.
3. Unzip the contents. Feel free to delete the .gitattributes and .gitignore files if you want to.
4. Right click on "updateSLAFileNames.ps1" and click "Run with PowerShell"
5. You'll see a popup for a few seconds of a blueish screen, that's just the code running. 
6. As long as your files have the correct name to begin with, you should see the filenames update automatically! 

The old files are moved to a Old_SLA_Files folder, feel free to do with them as you wish.