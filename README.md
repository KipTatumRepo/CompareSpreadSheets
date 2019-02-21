# CompareSpreadSheets
Python excel parser that will find the items in a spreadsheet by a list of given ItemId's

This was a quick parsing program that I put together to find an ItemId in an Excel spreadsheet with over 18,000 rows of varying length I called this file inHere.  I was given an Excel sheet with a list over 500 ItemId's the user wanted to find which I called 
findThese. 

The data is well formatted with a consistant layout where each cell represented a different data point and in numerical order by itemId.  The program will find the itemId in findThese and add it to an array.  The program then uses that array to search through the inHere file to find the itemId's that match.  When a match is found, that entire row is added to a new spreadsheet called Output.  

Since the data in the inHere file is in numerical order, the low end of the range to search is updated to start with the value of the itemId that was just found as there is never a need to go backward.  This will limit the size of the Excel sheet the program needs to search.

