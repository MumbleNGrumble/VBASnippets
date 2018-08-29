Attribute VB_Name = "ExcelVBAFunctions"
Option Explicit

Function CountChar(str As String, char As String, Optional caseSensitive As Boolean = False) As Long
    
    'Counts the occurance of the specified character in a string.
    'If caseSensitive = True, the Replace function will use vbBinaryCompare (0).
    'If caseSensitive = False, the Replace function will use vbTextCompare (1).
    '
    'Source: https://www.mrexcel.com/forum/excel-questions/234028-vba-count-substrings-string.html
    'str: String to search in.
    'char: Character to count. Must be a string of length one, otherwise incorrect results will be returned.
    'caseSensitive: True or False to take into account case sensitivity when counting the character.
    
    Dim caseSens As Byte
    If caseSensitive Then
        caseSens = 0
    Else
        caseSens = 1
    End If
    
    Dim baseCount As Long
    Dim endCount As Long
    
    baseCount = Len(str)
    endCount = Len(Replace(str, Find:=char, Replace:="", Start:=1, Count:=-1, Compare:=caseSens))
    
    CountChar = baseCount - endCount
    
End Function

Sub DeleteBlankRows(rng As Range, Optional includePartiallyBlankRows As Boolean = False)
    
    'Deletes blank rows in the specified range.
    'If includePartiallyBlankRows = False, it will delete rows that are blank across the entire worksheet.
    'If includePartiallyBlankRows = True, it will delete partially blank rows within the specified range.
    'Note: Excel will automatically determine the table boundaries in the range if includePartiallyBlankRows = True.
    'This means that if an entire column is blank in the range, it won't be taken into consideration for partial blanks.
    'Source: http://www.ozgrid.com/VBA/VBACode.htm
    '
    'rng: Range to search for blank rows
    'includePartiallyBlankRows: True or False to delete rows that are entirely blank or partially blank within the specified range.
    
    If Not includePartiallyBlankRows Then
        Dim i As Long
        For i = rng.Rows.Count To 1 Step -1
            If WorksheetFunction.CountA(rng.Rows(i)) = 0 Then
                rng.Rows(i).EntireRow.Delete
            End If
        Next
    Else
        rng.SpecialCells(xlBlanks).EntireRow.Delete
    End If
    
End Sub

Function FindLastColumn(ws As Worksheet, Optional rowArg As Long) As Long
    
    'Finds the last populated column in a worksheet and returns the column index.
    'If a row argument is specified, it finds the last populated column in that row.
    'Assumes no argument was provided if an invalid row was entered. Note: No error handling for row entries beyond the spreadsheet limit.
    'Returns 1 if the worksheet is empty.
    'Source: https://stackoverflow.com/a/11169920
    '
    'ws: Worksheet to search in.
    'rowArg: If specified, it will find the last populated column in that row.
    
    With ws
        If rowArg > 0 Then
            If Application.WorksheetFunction.CountA(.Rows(rowArg)) <> 0 Then
                FindLastColumn = .Rows(rowArg).Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
            Else
                FindLastColumn = 1
            End If
        Else
            If Application.WorksheetFunction.CountA(.Cells) <> 0 Then
                FindLastColumn = .Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
            Else
                FindLastColumn = 1
            End If
        End If
    End With
    
End Function

Function FindLastRow(ws As Worksheet, Optional col As String) As Long
    
    'Finds the last populated row in a worksheet and returns the row index.
    'If a column argument is specified, it finds the last populated row in that column.
    'Assumes no argument was provided if an invalid column was entered. Note: No error handling for column entries beyond the spreadsheet limit.
    'Returns 1 if the worksheet is empty.
    'Source: https://stackoverflow.com/a/11169920
    '
    'ws: Worksheet to search in.
    'col: If specified, it will find the last populated row in that column.
    
    With ws
        If Len(col) <= 3 And (col Like "[A-Z]" Or col Like "[a-z]") Then
            If Application.WorksheetFunction.CountA(.Columns(col)) <> 0 Then
                FindLastRow = .Columns(col).Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
            Else
                FindLastRow = 1
            End If
        Else
            If Application.WorksheetFunction.CountA(.Cells) <> 0 Then
                FindLastRow = .Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
            Else
                FindLastRow = 1
            End If
        End If
    End With
    
End Function

Sub GoFast(onOrOff As Boolean)
    
    'Turns off screen updating and alarms and sets calculations to manual.
    'Note: If your script errors out, the settings will stay as is. Be mindful of the calculation mode.
    'Source: https://stackoverflow.com/a/24936878
    '
    'onOrOff: True or False to enable or disable screen updating and alarms.
    
    Dim calcMode As XlCalculation
    calcMode = Application.Calculation
    
    With Application
        .ScreenUpdating = Not onOrOff
        .EnableEvents = Not onOrOff
        .DisplayAlerts = Not onOrOff
        
        If onOrOff Then
            .Calculation = xlCalculationManual
        Else
            .Calculation = calcMode
        End If
    End With
    
End Sub

Function RemoveNonAlphaNumeric(str As String, Optional removeSpaces As Boolean = False) As String
    
    'Removes non-alphanumeric characters from a string and returns the result as a string.
    'Source: https://www.extendoffice.com/documents/excel/651-excel-remove-non-numeric-characters.html
    '
    'str: String to remove non-alphanumeric characters from.
    
    Dim char As String
    Dim i As Integer
    Dim output As String
    output = ""
    
    For i = 1 To Len(str)
        char = Mid(str, i, 1)
        
        If char Like "[A-Z]" Or char Like "[a-z]" Or char Like "[0-9]" Then
            output = output & char
        End If
        
        If Not removeSpaces And char Like " " Then
            output = output & char
        End If
    Next
    
    RemoveNonAlphaNumeric = output
    
End Function

Function SheetExists(wsName As String, wb As Workbook) As Boolean
    
    'Checks if the specified name exists as a sheet.
    'Source: https://stackoverflow.com/a/6688482
    '
    'wsName: Name of the sheet to check for.
    'wb: Workbook to check in.
    
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = wb.Sheets(wsName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
    
End Function
