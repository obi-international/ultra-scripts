Sub FixNameSurnameCapitalization()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim splitParts As Variant
    Dim sheetName As String
    
    ' Prompt the user to enter the sheet name
    sheetName = InputBox("Enter the sheet name (e.g., original):", "Sheet Name", "e")
    
    ' Set the worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet not found!", vbExclamation
        Exit Sub
    End If
    
    ' Find the last row in column F
    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row ' Column F is the 6th column
    
    ' Loop through each row in column F starting from F3
    For i = 3 To lastRow ' Start from row 3
        cellValue = Trim(ws.Cells(i, 6).Value) ' Get value and remove extra spaces
        
        ' If the cell is not empty, fix the capitalization
        If Len(cellValue) > 0 Then
            ' Split the name into words
            splitParts = Split(cellValue, " ")
            
            ' Capitalize the first letter of each word
            Dim j As Integer
            For j = LBound(splitParts) To UBound(splitParts)
                splitParts(j) = Application.WorksheetFunction.Proper(LCase(splitParts(j)))
            Next j
            
            ' Rejoin words and update the cell
            ws.Cells(i, 6).Value = Join(splitParts, " ")
        End If
    Next i
    
    MsgBox "Capitalization fixed in Column F starting from F3.", vbInformation
End Sub
