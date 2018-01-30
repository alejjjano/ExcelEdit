import openpyxl
import string

filename = "CONSOLIDADOV3_TEST.xlsx"

# Load book and sheet
book = openpyxl.load_workbook(filename)
sheet = book.active

# Define column list
columns = [char for char in string.ascii_uppercase]

for i in string.ascii_uppercase:
    for j in string.ascii_uppercase:
        columns.append(i+j)


# Define start and stop points to edit
columnToStart = "B" # Guide value should be in this column
columnToEnd = "Z"

start = 4 # Label values should be in this column
stop = 1070

# Define list of columns
indexToStart = columns.index(columnToStart)
indexToEnd = columns.index(columnToEnd)

columnList = columns[indexToStart:indexToEnd+1]

# Initiate data structure to store



# Read row by row and store non-repeating data
for i in range(start, stop + 1):

#Write data in new book


# Save book to file
book.save("CONSOLIDADOV3_TEST_EDIT.xlsx")