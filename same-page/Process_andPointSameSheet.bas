Sub ProcessUnderScorePointInSameSheet()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim processedText As String
    Dim sheetName As String
    
    ' Prompt the user to enter the sheet name
    sheetName = InputBox("Enter the sheet name to edit (e.g., edited):", "Sheet Name", "edited")
    
    ' Set the worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet not found!"
        Exit Sub
    End If
    
    ' Find the last row in column H
    lastRow = ws.Cells(ws.Rows.Count, 8).End(xlUp).Row ' Column H is the 8th column
    
    ' Loop through each row in column H
    For i = 2 To lastRow ' Start from row 2 to skip headers
        cellValue = Trim(ws.Cells(i, 8).Value) ' Column H - Trim spaces
        
        ' Replace underscores `_` and points `.` with spaces
        cellValue = Replace(cellValue, "_", " ")
        cellValue = Replace(cellValue, ".", " ")
        
        ' Convert the processed text to Proper Case
        processedText = Application.WorksheetFunction.Proper(LCase(cellValue))
        
        ' Write the processed text to column I (9th column)
        ws.Cells(i, 9).Value = processedText
    Next i
    
    MsgBox "Processing complete. Modified values have been added to column I."
End Sub
