import openpyxl
import string


# Definir caracteres de columnas

columnChars = [char for char in string.ascii_uppercase]
for i in string.ascii_uppercase:
    for j in string.ascii_uppercase:
        columnChars.append(i+j)
print(columnChars)


def searchCol(columnName, sheet, header):
    """
    :param columnName: string, column to search
    :param sheet: sheet to look into
    :param header: the index of the header row
    :return: Index of column to search, as a letter
    """

    index = 0
    cursor = sheet[columnChars[index] + str(header)]
    while cursor.value is not None:
        if cursor.value == columnName:
            return columnChars[index]
        index += 1
        cursor = sheet[columnChars[index] + str(header)]
    raise ValueError("Column " + columnName + " not in Worksheet")


def searchRow(id, columnName, sheet, header):
    """
    :param columnName: the name of the column, as a string
    :param sheet: sheet to look at
    :param header: the index of header of the columns
    :param id: the value to look at in the column
    :return: the index of the row where id is found
    """
    columnIndex = searchCol(columnName, sheet, header)

    i = header + 1
    cursor = sheet[columnIndex+str(i)]
    while cursor.value is not None:
        if cursor.value == id:
            return i
        i += 1
        cursor = sheet[columnIndex + str(i)]
    raise ValueError("Value " + str(id) + " is not in column " + columnName)


def checkCol(sheet1, sheet2, columnToCheck, header):
    """
    :param sheet1: first sheet to check
    :param sheet2: second sheet to check
    :param columnToCheck: the column column to check, by name as a str
    :param header: the index of the header row
    :return: Boolean. True if the columns contain the same values in both sheets
    """

    columnIndex1 = searchCol(columnToCheck, sheet1, header)
    columnIndex2 = searchCol(columnToCheck, sheet2, header)

    # Track all values in sheet1
    listSheet1=[]
    i = header + 1
    cursor = sheet1[columnIndex1+str(i)]
    while cursor.value is not None:
        listSheet1.append(cursor.value)
        i += 1
        cursor = sheet1[columnIndex1 + str(i)]

    # Track all values in sheet2
    listSheet2=[]
    i = header + 1
    cursor = sheet2[columnIndex2+str(i)]
    while cursor.value is not None:
        listSheet2.append(cursor.value)
        i += 1
        cursor = sheet2[columnIndex2 + str(i)]

    #Check if items are repeated in both lists
    flag = True

    for item in listSheet1:
        if listSheet1.count(item) > 1:
            print("Repeated items in Sheet 1")
            flag = False

    for item in listSheet2:
        if listSheet2.count(item) > 1:
            print("Repeated items in Sheet 2")
            flag = False

    return sorted(listSheet1) == sorted(listSheet2) and flag


def columnList(sheet, columnName,header):
    """
    :param sheet: sheet to read
    :param columnName: column to read
    :param header: header of columns
    :return: list of all values in column
    """

    columnIndex = searchCol(columnName,sheet,header)
    colList = []

    i = header+1
    cursor = sheet[columnIndex + str(i)]
    while cursor.value is not None:
        colList.append(cursor.value)
        i += 1
        cursor = sheet[columnIndex + str(i)]

    return colList


def replaceValue(sheet1, sheet2, columnToEdit, tagColumn, id, header):
    """
    :param sheet1: sheet to replace values
    :param sheet2: sheet to get values from
    :param columnToEdit: columnToEdit, by Name
    :param tagColumn: name of the tag Column, as text
    :param id: id of the tag column to replace, wich is in both sheets
    :param header: the header of the columns
    :return: None, just replaces the value in the sheet1
    """

    rowIndexWrite = searchRow(id, tagColumn, sheet1, header)
    rowIndexRead = searchRow(id, tagColumn, sheet2, header)

    columnIndexWrite = searchCol(columnToEdit, sheet1, header)
    columnIndexRead = searchCol(columnToEdit, sheet2, header)

    sheet1[columnIndexWrite + str(rowIndexWrite)] = sheet2[columnIndexRead + str(rowIndexRead)].value

    print("Cell "+ columnIndexWrite + str(rowIndexWrite) + " updated to " + str(sheet2[columnIndexRead + str(rowIndexRead)].value))



# Leer Tablas C1 - C2

filenameC1 = "CONSOLIDADO_C21.xlsx"
filenameC2 = "CONSOLIDADO_C22.xlsx"

bookC1 = openpyxl.load_workbook(filenameC1)
bookC2 = openpyxl.load_workbook(filenameC2)

sheetC1 = bookC1.get_sheet_by_name("CONSOLIDADO")
sheetC2 = bookC2.get_sheet_by_name("CONSOLIDADO")

# Buscar columnas C1 - tag, A, B, C ... C2 - tag,  A, B, C ...
# Almacenar ubicación de columnas

tagColumn = "N°"

headerNum = 4

columnsToEdit = ["AREA RESPONSABLE DE INSPECCIÓN",
                 "RESPONSABLE DE INSPECCIÓN",
                 "FECHA DE INSPECCIÓN"]

# Revisar los valores de tag C1 y C2 para que:
# No se repitan y coincidan entre sí

if not checkCol(sheetC1, sheetC2, tagColumn, headerNum):
    raise ValueError("Values from " + tagColumn + " are repeated or not coherent in both tables")

# Reemplazar datos de C1 por los de C2 para
# las columnas A, B, C ...

tagColumnList = columnList(sheetC2,tagColumn,headerNum)

for column in columnsToEdit:
    for id in tagColumnList:
        replaceValue(sheetC1, sheetC2, column, tagColumn, id, headerNum)

# Guardar tabla con nuevo nombre

filenameCF = "CONSOLIDADO_CF2.xlsx"

bookC1.save(filenameCF)

print("Edited file saved as " + filenameCF)
