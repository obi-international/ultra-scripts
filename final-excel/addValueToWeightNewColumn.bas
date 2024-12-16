Sub AddValueToWeightColumnInNewColumn()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As Double
    Dim addedValue As Double
    Dim sheetName As String
    Dim startRow As Long
    Dim weightColumn As Integer
    Dim resultColumn As Integer
    
    ' Prompt the user to enter the sheet name
    sheetName = InputBox("Enter the sheet name (e.g., edited):", "Sheet Name", "edited")
    
    ' Set the worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet not found!"
        Exit Sub
    End If
    
    ' Prompt the user to enter the value to be added
    On Error Resume Next
    addedValue = CDbl(InputBox("Enter the value to add to the weights:", "Add Value"))
    On Error GoTo 0
    If IsNumeric(addedValue) = False Then
        MsgBox "Invalid number entered!"
        Exit Sub
    End If
    
    ' Set the starting point
    startRow = 19 ' Starting at row 19
    weightColumn = 11 ' Column K is the 11th column
    resultColumn = 12 ' Column L is the 12th column
    
    ' Find the last row in column K
    lastRow = ws.Cells(ws.Rows.Count, weightColumn).End(xlUp).Row
    
    ' Loop through each row in column K starting at row 19
    For i = startRow To lastRow
        ' Read the current weight value
        On Error Resume Next
        cellValue = CDbl(ws.Cells(i, weightColumn).Value)
        On Error GoTo 0
        
        ' Add the specified value and write the result to column L
        If IsNumeric(cellValue) Then
            ws.Cells(i, resultColumn).Value = cellValue + addedValue
        End If
    Next i
    
    MsgBox "Processing complete. The value " & addedValue & " has been added to the weights in column K and results saved in column L."
End Sub
