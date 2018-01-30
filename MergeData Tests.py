# Test search Col

# Case 1: List contains real columns

print("Test Case 1")

filenameC1 = "CONSOLIDADO_C1.xlsx"
bookC1 = openpyxl.load_workbook(filenameC1)
sheetC1 = bookC1.get_sheet_by_name("CONSOLIDADO")

headerNum = 4

columnsToEdit = ["CODIGO LOCAL",
                 "TOTAL MODULOS",
                 "ESTADO DE LOS TRABAJOS"]

indexes = []
for column in columnsToEdit:
    indexes.append(searchCol(column, sheetC1, headerNum))
    print(" Column ", column, "in index", searchCol(column, sheetC1, headerNum))


print(indexes)

# Case 2: List is empty

print("Test Case 2")

filenameC1 = "CONSOLIDADO_C1.xlsx"
bookC1 = openpyxl.load_workbook(filenameC1)
sheetC1 = bookC1.get_sheet_by_name("CONSOLIDADO")

headerNum = 4

columnsToEdit = []

indexes = []
for column in columnsToEdit:
    indexes.append(searchCol(column, sheetC1, headerNum))


print(indexes)

# Case 3: List contains non valid columns

print("Test Case 3")

filenameC1 = "CONSOLIDADO_C1.xlsx"
bookC1 = openpyxl.load_workbook(filenameC1)
sheetC1 = bookC1.get_sheet_by_name("CONSOLIDADO")

headerNum = 4

columnsToEdit = ["CODIGO LOCAL",
                 "TOTAL MODULOS",
                 "ESTADO DE LOS TRABAJOS",
                 "Hola"]

indexes = []
for column in columnsToEdit:
    try:
        indexes.append(searchCol(column, sheetC1, headerNum))
        print(" Column ", column, "in index", searchCol(column, sheetC1, headerNum))
    except:
        print("Test 3 successfull")

print(indexes)


# Case 4: List contains accented text

print("Test Case 4")

filenameC1 = "CONSOLIDADO_C1.xlsx"
bookC1 = openpyxl.load_workbook(filenameC1)
sheetC1 = bookC1.get_sheet_by_name("CONSOLIDADO")

headerNum = 4

columnsToEdit = ["TIPO DE INSPECCIÓN"]

indexes = []
for column in columnsToEdit:
    indexes.append(searchCol(column, sheetC1, headerNum))
    print(" Column ", column, "in index", searchCol(column, sheetC1, headerNum))

print(indexes)






# Test cases for checkCol

# Test 1: When both columns have equal values

sheet1 = openpyxl.load_workbook("TESTS.xlsx").get_sheet_by_name("T1-1")
sheet2 = openpyxl.load_workbook("TESTS.xlsx").get_sheet_by_name("T1-2")
print(checkCol(sheet1,sheet2,"COLUMNA",1))

# Test 2: When both columns have different values

sheet1 = openpyxl.load_workbook("TESTS.xlsx").get_sheet_by_name("T2-1")
sheet2 = openpyxl.load_workbook("TESTS.xlsx").get_sheet_by_name("T2-2")
print(checkCol(sheet1,sheet2,"COLUMNA",1))

# Test 3: When one column is empty

sheet1 = openpyxl.load_workbook("TESTS.xlsx").get_sheet_by_name("T3-1")
sheet2 = openpyxl.load_workbook("TESTS.xlsx").get_sheet_by_name("T3-2")
print(checkCol(sheet1,sheet2,"COLUMNA",1))

# Test 4: When both columns are empty

sheet1 = openpyxl.load_workbook("TESTS.xlsx").get_sheet_by_name("T4-1")
sheet2 = openpyxl.load_workbook("TESTS.xlsx").get_sheet_by_name("T4-2")
print(checkCol(sheet1,sheet2,"COLUMNA",1))





# Test cases for searchRow

# Test 1: value in column

sheet = openpyxl.load_workbook("TESTS.xlsx").get_sheet_by_name("T1-1")
if searchRow(20, "COLUMNA", sheet, 1) == 21:
    print("Test 1 successfull")
else:
    print("Test 1 wrong")

# Test 2: value not in column

sheet = openpyxl.load_workbook("TESTS.xlsx").get_sheet_by_name("T1-1")
try:
    print(searchRow(50, "COLUMNA", sheet, 1))
except ValueError:
    print("Test 2 succesfull")
else:
    print("Test 2 wrong")

# Test 3: value multiple times in column

sheet = openpyxl.load_workbook("TESTS.xlsx").get_sheet_by_name("T5-B")
if searchRow(20, "COLUMNA", sheet, 1) == 21:
    print("Test 3 successfull")
else:
    print("Test 3 wrong")

# Test 4: Empty column

sheet = openpyxl.load_workbook("TESTS.xlsx").get_sheet_by_name("T4-2")
try:
    print(searchRow(50, "COLUMNA", sheet, 1))
except ValueError:
    print("Test 4 succesfull")
else:
    print("Test 4 wrong")