# excelVBAMacros
Written for CDC Summer 2019

All macros are written in Visual Basic.

# Contents

receiptParser.txt - Taking the existing Outlook session, looks for a specific purchase receipt email in the "Orders" subfolder of the inbox.
Parses the file and writes each of the fields into the same row in an excel entry. Always writes to the end of the file.

formatRunLog.txt - Taking an existing log of a specific format, formats the columns, sorts by date, and cleans up some of the data columns with incorrect values.
Provides user with a prompt to mark some values as "Yes" based on the value in another column.
Log must be formatted in the correct format for the macro to work.

JeremyMacros.xlam - Excel add-in with both the previous macros implemented.
Usage: 
1. Open up a session of Excel. 
2. Select File > Options > Add-ins. Near the bottom, there is a section that says "Manage: " with a drop-down menu. Make sure the menu says Excel add-ins.
3. Click "Go" on the button next to the menu.
4. In the resulting window, click "Browse" on the right hand side.
5. Find JeremyMacros.xlam in where the file was downloaded. Double-click the file.
6. Answer "Yes" when asked if the file should be copied to the network drive/wherever Excel is.
7. JeremyMacros should show up with a check in the box on the left. If the box is not checked, make sure to check the box. Press "Ok" when done to close the window.
8. In order to add these macros to the top bar, go to File > Options > Customize Ribbon. Click "New Tab".
9. On the right hand side, highlight the "New Group" label that appears. Then, on the left hand side, click the drop down menu and find "Macros".
10. Click on a macro that you would like to use, then click the "Add" button in the middle of the two panes.
11. Repeat step 10 for each macro you would like to add. Click "Ok" when you're done.
12. Return to the excel sheet and press Alt +  F11. This will open up the VBA Editor. 
13. On the top bar, find "Tools" then select "References".
14. Scroll down until you find the references that start with "Microsoft Outlook". Check each of these references. This is so the VBA code will be able to access your Outlook session.
15. Press "Ok" when completed. Step 12, 13, and 14 have to be completed for each new excel file that needs the receipt parser macro.
16. Done! All of the macros should be able to be used to their full extent!
