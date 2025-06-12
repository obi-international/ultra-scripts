Sub FixColumnCapitalizationFromTop()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim splitParts As Variant
    Dim sheetName As String
    Dim colLetter As String
    Dim colNumber As Long
    
    ' Prompt for sheet name
    sheetName = InputBox("Enter the sheet name (e.g., e):", "Sheet Name", "e")
    If sheetName = "" Then Exit Sub
    
    ' Prompt for column letter
    colLetter = InputBox("Enter the column letter to fix (e.g., F):", "Column Letter", "F")
    If colLetter = "" Then Exit Sub
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet not found!", vbExclamation
        Exit Sub
    End If

    ' Validate and convert column letter to number safely
    On Error Resume Next
    colNumber = Range(colLetter & "1").Column
    If Err.Number <> 0 Or colNumber = 0 Then
        MsgBox "Invalid column letter!", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Find last row
    lastRow = ws.Cells(ws.Rows.Count, colNumber).End(xlUp).Row
    
    ' Loop from row 1
    For i = 2 To lastRow
        cellValue = Trim(ws.Cells(i, colNumber).value)
        
        If Len(cellValue) > 0 Then
            splitParts = Split(cellValue, " ")
            Dim j As Integer
            For j = LBound(splitParts) To UBound(splitParts)
                splitParts(j) = Application.WorksheetFunction.Proper(LCase(splitParts(j)))
            Next j
            ws.Cells(i, colNumber).value = Join(splitParts, " ")
        End If
    Next i
    
    MsgBox "Capitalization fixed in Column " & colLetter & ".", vbInformation
End Sub

