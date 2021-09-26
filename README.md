# MergeTablesVBA
Create a VB script to load a workbook, and read every sheet to merge in a table. 

First, create an array of all tabs starting with zero ("0"). All other tabs (like "Cover") are irrelevant.

For each data tab in this array

First, just copy the first tab to the Merger tab, as obviously everything needs to come over.
For each data tab starting at tab 2 to the tab before the Merger tab

Identify the primary key column (ex. Customer ID)

For each row on current data tab
get the value at current row and primary key column
attempt to locate the row with the primary key on the merger tab

if the primary key exists, use that row
else, use the next available row on the merger tab and paste the primary key
EndIF

For each column on the data tab

attempt to find a matching column on the merger tab

if matching column exists, use that column
else, add column header to new column and use that column
EndIF

Paste value from datatab current row/current column to merger tab primary row, relevant column

End For Each #column for current row
End For Each #Row on current data tab
End For Each #data tab
End For Each #array
