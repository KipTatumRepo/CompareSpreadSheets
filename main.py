import xlrd
import xlwt
import getpass

# this gets login name of the current computer user
userName = getpass.getuser()
print("Opening and reading the first spreadsheet")
# where and what files to open, I know they are on the Desktop findThese.xlsx is a spreadsheet that contains a table of
# desired itemId's
# InHere.xlsx is a spreadsheet that contains all items
pathopen = "C:\\Users\\" + userName + "\\Desktop\\findThese.xlsx"
pathfulllistopen = "C:\\Users\\" + userName + "\\Desktop\\InHere.xlsx"

# create 2 workbooks
WorkBookRead = xlrd.open_workbook(pathopen)
OtherWorkBook = xlrd.open_workbook(pathfulllistopen)

# add sheets to those 2 workbooks
WorkSheetRead = WorkBookRead.sheet_by_index(0)
OtherWorkSheetRead = OtherWorkBook.sheet_by_index(0)

# create a workbook to be saved on the users desktop, add a sheet and name that sheet All in List
WorkBookWrite = xlwt.Workbook("C:\\Users\\" + userName + "\\Desktop\\output.xlsx")
WorkSheetWrite = WorkBookWrite.add_sheet('All in List')

# -------------------------------------------------------------------------------#
#  this section is where we read the 2 spread sheets, add their ItemId's to      #
#  array's so we can compare them to find Items that are the same in both sheets #
# -------------------------------------------------------------------------------#

# initialize 3 arrays, 2 to hold the ItemId from each file, 1 to hold the Id's from the
# all items workbook that we are looking for
IdList = []
OtherIdList = []
SameId = []

# add each ItemId from the the workbook that has the itemId's we want to find
for rows in range(WorkSheetRead.nrows):
    ItemId = WorkSheetRead.cell(rows, 1).value
    IdList.append(ItemId)
    if rows == 200 or rows == 350 or rows == 500:
        print(".")

print("Opening and reading the second spreadsheet")
# add each of the ItemId from all item workbook provided
for rows in range(OtherWorkSheetRead.nrows):
    OtherItemId = OtherWorkSheetRead.cell(rows, 0).value
    OtherIdList.append(OtherItemId)
    if rows == 0 or rows == 3000 or rows == 6000 or rows == 9000 or rows == 12000 or rows == 15000 or rows == 18000:
        print(".")
# remove the first row(it's a header)
# IdList.pop(0)

# this is where we compare 2 lists to find the itemId we are looking for in the all item list
# if a match is found add that ItemId to the SameId array
print("Comparing the 2 files to find matches")

# start = 0
LowRange = 0
for i in range(len(IdList)):
    found = False
    IdListItem = int(IdList[i])
    for j in range(LowRange, len(OtherIdList)):
        # if the ItemId is found in the inHere file, set the row to start the next loop at and exit the loop
        if found:
            LowRange = j
            break
        OtherIdListItem = int(OtherIdList[j])
        if IdListItem == OtherIdListItem:
            SameId.append(OtherIdList[j])
            found = True

    if i == 200 or i == 350 or i == 500:
        print(".")
# print(SameId)

print("File comparison complete")
# ----------------------------------------------------------------------------#
#  this section is where the data is saved to another excel book to be viewed #
# ----------------------------------------------------------------------------#

# the path to save the file(on the users desktop)
newfolderpath = "C:\\Users\\" + userName + "\\Desktop\\Output.xls"

# initialize an array that is 1 row from the all item spreadsheet
oneRow = []
rowCount = 0
for i in range(len(SameId)):
    # get an Id from SameId Array
    NewIdListItem = SameId[i]

    # create the row we want to write in the new spreadsheet
    row = WorkSheetWrite.row(i)

    # iterate over each row in the all item spreadsheet to see if there is a ItemId that matches the
    # current ItemId from the SameId Array
    start = 0
    nextDot = 1
    found = False
    for rows in range(start, OtherWorkSheetRead.nrows):
        # No need to continue iterating over the list if the ItemId is found
        if found:
            start = rows
            break

        # if there is a match add all the cells from that row to the oneRow array
        # the rows in the all item spreadsheet are not of a fixed length
        if NewIdListItem == OtherWorkSheetRead.cell(rows, 0).value:  # changed to zero maybe back to 1
            for j in range(OtherWorkSheetRead.row_len(rows)):
                oneRow.append(OtherWorkSheetRead.cell(rows, j).value)
            start = rows
            found = True

            # when all the cells from the current row have been added to the oneRow array
            # for the length of that array(because each row will have a different length) write each cell to
            # the new spreadsheet
            for l in range(len(oneRow)):
                row.write(l, oneRow[l])
            rowCount += 1
            print("Writing row " + rowCount.__str__() + " to output.xls on your desktop")

    # Here we clear the oneRow array to create a "new" row
    oneRow.clear()

# save the workbook on the users desktop
WorkBookWrite.save(newfolderpath)
print("The new spreadsheet is on your desktop, this program is finished running")