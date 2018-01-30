import openpyxl


filename = "CONSOLIDADOV3.xlsx"

# Load book and sheet
book = openpyxl.load_workbook(filename)
sheet = book.active

# Define start and stop points to edit
columnToRead = "Y"
start = 5
stop = 1070

# Define column to write
columnToWrite = "AA"

# Start counter
counter = 0

# Read every cell in range, by rows
for i in range(start, stop+1):
    # Evaluates if the value has changed
    if sheet[columnToRead+str(i)].value != sheet[columnToRead+str(i-1)].value:
        counter += 1
    # Write counter value in column to write
    sheet[columnToWrite+str(i)] = counter
    print("Cell "+columnToWrite+str(i)+" updated to "+str(counter))


book.save("CONSOLIDADOV3_EDIT.xlsx")
