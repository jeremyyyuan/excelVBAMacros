# excelVBAMacros
Written for CDC Summer 2019

All macros are written in Visual Basic.

# Contents

receiptParser.txt - Taking the existing Outlook session, looks for a specific purchase receipt email in the "Orders" subfolder of the inbox.
Parses the file and writes each of the fields into the same row in an excel entry. Always writes to the end of the file.

formatRunLog.txt - Taking an existing log of a specific format, formats the columns, sorts by date, and cleans up some of the data columns with incorrect values.
Provides user with a prompt to mark some values as "Yes" based on the value in another column.
Log must be formatted in the correct format for the macro to work.
