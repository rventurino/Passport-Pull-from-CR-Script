Sub pullFromCRBackupPassportMAIN()
ProgramStartMessage
GenerateBlankPPWorksheets
ParsePPFromXML
DepartmentResolve
ScriptComplete
End Sub


Sub GenerateBlankPPWorksheets()
Dim rootFilePath As String
rootFilePath = ThisWorkbook.Path
'================================================================Parse XML Files from Backup to Excel Files
'Disable alerts from Excel
    'Prevents Excel from prompting for "Save information in clipboard" when closing large Passport Reference Table Excel Sheets
                Application.DisplayAlerts = False

'Create the new Excel Sheet
 ' Create a new workbook
    Dim wb As Workbook
    Set wb = Workbooks.Add

    ' Save the workbook
    wb.SaveAs FileName:= rootFilePath & "\New Microsoft Excel Worksheet.xlsx"

    ' Open the workbook
    Workbooks.Open (rootFilePath & "\New Microsoft Excel Worksheet.xlsx")
'Generate Excel worksheets, color them, and make the headings for each worksheet (Tax, Items, Departments, etc.)

'rename first sheet to tax
Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Sheet1").Name = "Tax Codes"

'Add Sheets and Name them
Sheets.Add(After:=Sheets("Tax Codes")).Name = "Categories"
Sheets.Add(After:=Sheets("Categories")).Name = "Departments"
Sheets.Add(After:=Sheets("Departments")).Name = "Department Resolve"
Sheets.Add(After:=Sheets("Department Resolve")).Name = "Items"
Sheets.Add(After:=Sheets("Items")).Name = "Items Cleaned"
Sheets.Add(After:=Sheets("Items Cleaned")).Name = "Items XREF"
Sheets.Add(After:=Sheets("Items XREF")).Name = "Linked Items"
Sheets.Add(After:=Sheets("Linked Items")).Name = "Invalid SKU"
Sheets.Add(After:=Sheets("Invalid SKU")).Name = "Duplicate SKU"

'Update Color Pallette for Tabs
Sheets("Tax Codes").Tab.Color = RGB(255, 192, 0)
Sheets("Categories").Tab.Color = RGB(0, 176, 80)
Sheets("Departments").Tab.Color = RGB(146, 208, 80)
Sheets("Department Resolve").Tab.Color = RGB(146, 208, 80)
Sheets("Items").Tab.Color = RGB(0, 176, 240)
Sheets("Items Cleaned").Tab.Color = RGB(0, 176, 240)
Sheets("Items XREF").Tab.Color = RGB(0, 112, 192)
Sheets("Linked Items").Tab.Color = RGB(0, 32, 96)
Sheets("Invalid SKU").Tab.Color = RGB(255, 0, 0)
Sheets("Duplicate SKU").Tab.Color = RGB(255, 0, 0)

'Tax Headings
Sheets("Tax Codes").Range("A1").Value = "Tax Number"
Sheets("Tax Codes").Range("B1").Value = "Tax Name"
Sheets("Tax Codes").Range("C1").Value = "CSO Tax 1"
Sheets("Tax Codes").Range("D1").Value = "CSO Tax 2"
Sheets("Tax Codes").Range("E1").Value = "CSO Tax 3"
Sheets("Tax Codes").Range("F1").Value = "CSO Tax 4"

'Categories Headings
Sheets("Categories").Range("A1").Value = "Category Name"
Sheets("Categories").Range("B1").Value = "Category Number"

'Departments Headings
Sheets("Departments").Range("A1").Value = "Department Number"
Sheets("Departments").Range("B1").Value = "Department Name"
Sheets("Departments").Range("C1").Value = "Negative Department"


'Items Headings
Sheets("Items").Range("A1").Value = "UPC"
Sheets("Items").Range("B1").Value = "UPC Type"
Sheets("Items").Range("C1").Value = "Taxation"
Sheets("Items").Range("D1").Value = "Item Description"
Sheets("Items").Range("E1").Value = "CR Description"
Sheets("Items").Range("F1").Value = "Department"
Sheets("Items").Range("G1").Value = "Age Restriction"
Sheets("Items").Range("H1").Value = "Food Stamps/EBT"
Sheets("Items").Range("I1").Value = "Unit Price"
Sheets("Items").Range("J1").Value = "Return Price"
Sheets("Items").Range("K1").Value = "Retail Price"
Sheets("Items").Range("L1").Value = "Linked Items"
Sheets("Items").Range("M1").Value = "Product Code"

'Items XREF Headings
Sheets("Items XREF").Range("A1").Value = "New PLU"
Sheets("Items XREF").Range("B1").Value = "Existing PLU"

'Items Cleaned 
Sheets("Items Cleaned").Range("A1").Value = "UPC"
Sheets("Items Cleaned").Range("B1").Value = "UPC Type"
Sheets("Items Cleaned").Range("C1").Value = "Item Description"
Sheets("Items Cleaned").Range("D1").Value = "CR Description"
Sheets("Items Cleaned").Range("E1").Value = "Department"

'Items Linked Items Headings
Sheets("Linked Items").Range("A1").Value = "Item UPC"
Sheets("Linked Items").Range("B1").Value = "Linkable Item UPC"

'Invalid SKU Headings
Sheets("Items Cleaned").rows(1).Copy
Sheets("Invalid SKU").Range("A1").PasteSpecial Paste:=xlPasteValues

'Duplicate SKU Headings
Sheets("Items Cleaned").rows(1).Copy
Sheets("Duplicate SKU").Range("A1").PasteSpecial Paste:=xlPasteValues

End Sub
Sub ParsePPFromXML()
'========================================================================================================================================
'========================================================================================================================================
'declare file location variables
Dim rootFilePath As String
Dim filePath As String
rootFilePath = ThisWorkbook.Path

'Convert tax to XLSX
    'set xmlFilePath to location of tax
    xmlFilePath = rootFilePath & "\Backup\ReferenceTables\GlobalSTORE_TAXABILITY.XML"
    'load File in Excel
    Workbooks.OpenXML FileName:=xmlFilePath, LoadOption:=xlXmlLoadImportToList
    'Save and close, set file name
    ActiveWorkbook.Close Savechanges:=True, FileName:=rootFilePath & "\GlobalSTORE_TAXABILITY.XML.xlsx"

'Convert Departments to XLSX
    'set xmlFilePath to location of tax
    xmlFilePath = rootFilePath & "\Backup\ReferenceTables\GlobalSTORE_STORE_DEPARTMENT.XML"
    'load File in Excel
    Workbooks.OpenXML FileName:=xmlFilePath, LoadOption:=xlXmlLoadImportToList
    'Save and close, set file name
    ActiveWorkbook.Close Savechanges:=True, FileName:=rootFilePath & "\GlobalSTORE_STORE_DEPARTMENT.XML.xlsx"

'Convert Items to XLSX ================================================================================================================================
Application.DisplayAlerts = True

    'set xmlFilePath to location of tax
    xmlFilePath = vbNullString
    xmlFilePath = rootFilePath & "\Backup\ReferenceTables\GlobalSTORE_PLU.XML"
    'load File in Excel
    Workbooks.OpenXML FileName:=xmlFilePath, LoadOption:=xlXmlLoadImportToList
    'Save and close, set file name
    ActiveWorkbook.Close Savechanges:=True, FileName:=rootFilePath & "\GlobalSTORE_PLU.XML.xlsx"
  Application.DisplayAlerts = False   
'Convert Items XREF to XLSX
    'set xmlFilePath to location of tax
    xmlFilePath = rootFilePath & "\Backup\ReferenceTables\GlobalSTORE_ITEM_XREF.XML"
    'load File in Excel
    Workbooks.OpenXML FileName:=xmlFilePath, LoadOption:=xlXmlLoadImportToList
    'Save and close, set file name
    ActiveWorkbook.Close Savechanges:=True, FileName:=rootFilePath & "\GlobalSTORE_ITEM_XREF.XML.xlsx"

'Convert Linked Items to XLSX
    'set xmlFilePath to location of tax
    xmlFilePath = rootFilePath & "\Backup\ReferenceTables\GlobalSTORE_PLU_GROUP.XML"
    'load File in Excel
    Workbooks.OpenXML FileName:=xmlFilePath, LoadOption:=xlXmlLoadImportToList
    'Save and close, set file name
    ActiveWorkbook.Close Savechanges:=True, FileName:=rootFilePath & "\GlobalSTORE_PLU_GROUP.XML.xlsx"


'=====================================================================================================================================================
'=====================================================================================================================================================
'=====================================================================================================================================================
'=====================================================================================================================================================

'Parse data from each seperate XML converted sheet in to the New Excel Sheet
'YOURLISTOBJECT.HeaderRowRange.Cells.Find("A_VALUE").Column


'Copy Taxation to Generated Spreadsheet
    'Open Desired Workbook
    Workbooks.Open (ThisWorkbook.Path & "\GlobalSTORE_TAXABILITY.XML.xlsx")
        'Use searchByHeader method to copy data to tax page by header name
        searchByHeader "TAXBLTY_CD", "GlobalSTORE_TAXABILITY.XML.xlsx", "Tax Codes", "A2"
        searchByHeader "DESCR", "GlobalSTORE_TAXABILITY.XML.xlsx", "Tax Codes", "B2"
                 'fit UPC column to contents
                Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Tax Codes").columns("A:Z").AutoFit
                    Workbooks("GlobalSTORE_TAXABILITY.XML.xlsx").Close Savechanges:=False



'Copy Departments to Generated Spreadsheet
    'Open Desired Workbook
    Workbooks.Open (ThisWorkbook.Path & "\GlobalSTORE_STORE_DEPARTMENT.XML.xlsx")

        'Use searchByHeader method to copy data to department page by header name
        searchByHeader "STR_HIER_ID", "GlobalSTORE_STORE_DEPARTMENT.XML.xlsx", "Departments", "A2"
        searchByHeader "DESCR", "GlobalSTORE_STORE_DEPARTMENT.XML.xlsx", "Departments", "B2"
        searchByHeader "CS_NEG_DEPT_FG", "GlobalSTORE_STORE_DEPARTMENT.XML.xlsx", "Departments", "C2"

            'Generate 1:1 Categories based on dept list
            Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Departments").Range("A2:B10000").Copy
            Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Categories").Range("A2").PasteSpecial Paste:=xlPasteValues

            'fit column to contents
                Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Categories").columns("A:Z").AutoFit
            'fit column to contents
                Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Departments").columns("A:Z").AutoFit
            Workbooks("GlobalSTORE_STORE_DEPARTMENT.XML.xlsx").Close Savechanges:=False         
'Copy Items to Generated Spreadsheet
'Open Desired Workbook
    Workbooks.Open (ThisWorkbook.Path & "\GlobalSTORE_PLU.XML.xlsx")
        'Copy and Paste UPCs as Values
        searchByHeader "PLU_ID", "GlobalSTORE_PLU.XML.xlsx", "Items", "A2"
        'Copy and Paste Taxation, Item Description, CR Description, Department as Values
        'taxation
        searchByHeader "TAXBLTY_CD", "GlobalSTORE_PLU.XML.xlsx", "Items", "C2"
        'Item Description
        searchByHeader "DSPL_DESCR", "GlobalSTORE_PLU.XML.xlsx", "Items", "D2"
        'Item CR Description
        searchByHeader "RCPT_DESCR", "GlobalSTORE_PLU.XML.xlsx", "Items", "E2"
        'Department
        searchByHeader "STR_HIER_ID", "GlobalSTORE_PLU.XML.xlsx", "Items", "F2"
        'Copy and Paste Age Restriction
        searchByHeader "SLS_RSTRCT_GRP", "GlobalSTORE_PLU.XML.xlsx", "Items", "G2"
        'Copy and Paste Food Stamps
        searchByHeader "FD_STMP_FG", "GlobalSTORE_PLU.XML.xlsx", "Items", "H2"
        'Copy and Paste Unit Price
        searchByHeader "UNT_PRC", "GlobalSTORE_PLU.XML.xlsx", "Items", "I2"
        'Copy and Paste Return Price
        searchByHeader "RTN_PRC", "GlobalSTORE_PLU.XML.xlsx", "Items", "J2"
        'Copy and Paste Linked Item Flag
        searchByHeader "LINK_ITM_FG", "GlobalSTORE_PLU.XML.xlsx", "Items", "L2"
        'Copy Product Code Number
        searchByHeader "CS_NET_PROD_CD", "GlobalSTORE_PLU.XML.xlsx", "Items", "M2"

       'Ensure correct page (Items)
        Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Items").Activate
            'Format to number format for visibility
            With Range("A2:A100000")
            .NumberFormat = "0"
            .Value = .Value
            End With
                'fit UPC column to contents
                Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Items").columns("A:Z").AutoFit
            'sort by column A
            Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Items").columns("A:M").Sort Key1:=Range("A:A"), Order1:=xlAscending, Header:=xlYes
            'remove non numeric UPC rows
            DeleteNonNumericRows
            Workbooks("GlobalSTORE_PLU.XML.xlsx").Close Savechanges:=False
'Copy Items XREF to Generated Spreadsheet
'Open Desired Workbook
    Workbooks.Open (ThisWorkbook.Path & "\GlobalSTORE_ITEM_XREF.XML.xlsx")
        'Copy and Paste UPCs as Values
        searchByHeader "SCAN_ID", "GlobalSTORE_ITEM_XREF.XML.xlsx", "Items XREF", "A2"
        searchByHeader "PLU_ID", "GlobalSTORE_ITEM_XREF.XML.xlsx", "Items XREF", "B2"
        'Ensure correct page (Items)
        Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Items XREF").Activate
            'Format to number format for visibility
            With Range("A2:B100000")
            .NumberFormat = "0"
            .Value = .Value
            End With
        'Sort Ascending
        Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Items XREF").columns("A:B").Sort Key1:=Range("A:A"), Order1:=xlAscending, Header:=xlYes
        'Delete rows if the old and new UPC are the same
        ItemXREFDeleteRowsIfEqual
        'Copy headers
        Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Items").Range("B1:M1").Copy
        Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Items XREF").Range("C1").PasteSpecial Paste:=xlPasteValues
        Workbooks("GlobalSTORE_ITEM_XREF.XML.xlsx").Close Savechanges:=False        
'Check linked items list, import if they exist, continue onward if they don't
LinkedItemsWithErrorHandling

End Sub

Sub DepartmentResolve()


'Copy Data to Department Resolve Sheet

'Generate Department Resolve Sheet
Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Items").columns("F").Copy
Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Department Resolve").Range("A1").PasteSpecial Paste:=xlPasteValues
Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Department Resolve").Range("B1").Value = "Department Name"
Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Department Resolve").Range("C1").Value = "Negative Department"
Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Items").columns("C").Copy
Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Department Resolve").Range("D1").PasteSpecial Paste:=xlPasteValues
Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Items").Range("G1:H100000").Copy
Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Department Resolve").Range("E1").PasteSpecial Paste:=xlPasteValues
Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Items").columns("M").Copy
Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Department Resolve").Range("G1").PasteSpecial Paste:=xlPasteValues


'Aggregate Data: Department Resolve Sheet
'Remove Duplicates
Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Department Resolve").Activate
Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Department Resolve").UsedRange.RemoveDuplicates columns:=Array(1, 2, 3, 4, 5, 6, 7), Header:=xlYes
'Sort
Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Department Resolve").Sort.SortFields.Clear
Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Department Resolve").UsedRange.Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
'VLOOKUP Negative and Name
Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Department Resolve").Range("B2", Range("A2").End(xlDown).Offset(0, 1)).Formula = "=VLOOKUP(A2,Departments!$A:$B,2,FALSE)"
Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Department Resolve").Range("C2", Range("A2").End(xlDown).Offset(0, 2)).Formula = "=VLOOKUP(A2,Departments!$A:$C,3,FALSE)"
'Append Department page data to bottom of Department Resolve
Dim lastRow As String

lastRow = ActiveSheet.Cells(rows.Count, "B").End(xlUp).row + 1
Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Departments").Range("A2:C10000").SpecialCells(xlCellTypeConstants).Copy
Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Department Resolve").Range("A" & lastRow).PasteSpecial Paste:=xlPasteValues
HighlightDuplicates
'fit UPC column to contents
Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets("Department Resolve").columns("A:Z").AutoFit
DepartmentResolveDupesMessage
End Sub


Sub ScriptComplete()
Application.DisplayAlerts = True
MsgBox Prompt:="Excel Import from Passport Complete ", Buttons:=vbOKOnly, Title:="MsgBox"


End Sub

Sub HighlightDuplicates()
    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range
    Dim i As Long
    
    'Change the column letter to the column you want to check for duplicates
    Set rng = Range("A1:A" & Cells(rows.Count, "A").End(xlUp).row)
    
    lastRow = rng.Cells.Count
    
    For i = 1 To lastRow
        Set cell = rng.Cells(i, 1)
        
        'Check if the value is a duplicate
        If WorksheetFunction.CountIf(rng, cell.Value) > 1 Then
            cell.Interior.Color = RGB(230, 184, 183)
            cell.Font.Color = RGB(99, 37, 35)
        Else
            cell.Interior.ColorIndex = xlNone
        End If
    Next i
End Sub

Sub DepartmentResolveDupesMessage()
MsgBox Prompt:="At the Bottom of the Department Resolve list, remove all HIGHLIGHTED departments with no values beyond column C", Buttons:=vbOKOnly, Title:="MsgBox"

End Sub

Sub ProgramStartMessage()

MsgBox Prompt:="This program will import the Passport Backup to an Excel sheet. Please close all non-related Excel Windows. Please do not touch Excel until the next dialog box appears", Buttons:=vbOKOnly, Title:="MsgBox"

End Sub
Sub LinkedItemsWithErrorHandling()
    'Ignore error which occurs when the linked items reference table is empty.
    On Error Resume Next
    
    'Copy Linked Items to Generated Spreadsheet
    'Open Desired Workbook
    Workbooks.Open (ThisWorkbook.Path & "\GlobalSTORE_PLU_GROUP.XML.xlsx")
        'Copy and Paste UPCs as Values
        searchByHeader "PLU_ID", "GlobalSTORE_PLU_GROUP.XML.xlsx", "Linked Items", "A2"
        searchByHeader "COMPT_PLU_ID", "GlobalSTORE_PLU_GROUP.XML.xlsx", "Linked Items", "B2"

            'Close
            Workbooks("GlobalSTORE_PLU_GROUP.XML.xlsx").Close Savechanges:=False

End Sub

Sub searchByHeader(ByVal headerQuery As String, ByVal sheetToSearch As String, ByVal destPageName As String, ByVal destPasteIndex)
    'HOW TO USE THIS SubRoutine:
        ' call the subroutine and pass it the following four variables
            'headerQuery: The name of the header to search
            'sheetToSearch: The (already opened) workbook you will be searching
            'detPageName: The name of the page you will be PASTING the data to
            'destPasteIndex: The Row and column you will PASTE TO, for example B2
    Dim HeaderName As String
    Dim ColumnIndex As Integer
    Dim CopyRange As Range
    
    Workbooks(sheetToSearch).Worksheets(1).Activate
    
    ' Enter the header name of the column to be copied
    HeaderName = headerQuery
    
    ' Search for the header name in the first row of the active worksheet
    ColumnIndex = WorksheetFunction.Match(HeaderName, ActiveSheet.rows(1), 0)
    
    ' Check if the header name was found
    If Not IsError(ColumnIndex) Then
        ' Select the column based on the header name and the last row of the active worksheet
        Set CopyRange = Range(Cells(2, ColumnIndex), Cells(rows.Count, ColumnIndex).End(xlUp))
        ' Copy the selected column to the clipboard
        CopyRange.SpecialCells(xlCellTypeConstants).Copy
    Else
        MsgBox "Header name not found!", vbExclamation, "Error"
    End If
    
    Workbooks("New Microsoft Excel Worksheet.xlsx").Worksheets(destPageName).Range(destPasteIndex).PasteSpecial Paste:=xlPasteValues
End Sub

Sub DeleteNonNumericRows()
    Dim lastRow As Long
    Dim i As Long
    
    'Get the last row in column A
    lastRow = Cells(rows.Count, "A").End(xlUp).row
    
    'Loop through each cell in column A
    For i = lastRow To 2 Step -1
        'Check if the value is not numeric
        If Not IsNumeric(Cells(i, "A").Value) Then
            'Delete the entire row if it's not numeric
            rows(i).Delete
        End If
    Next i
End Sub

Sub ItemXREFDeleteRowsIfEqual()
    
    Dim i As Integer
    Dim lastRow As Long
    
    lastRow = Cells(rows.Count, "A").End(xlUp).row 'get the last row in column A
    
    For i = lastRow To 2 Step -1 'loop through each row in column A in reverse order
        If Cells(i, 1).Value = Cells(i, 2).Value Then 'compare the values in the two cells
            rows(i).Delete 'if the values are equal, delete the entire row
        End If
    Next i

End Sub