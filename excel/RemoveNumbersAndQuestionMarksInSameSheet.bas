Sub RemoveNRandSymbols()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim processedText As String
    Dim sheetName As String
    
    ' Prompt the user to enter the sheet name
    sheetName = InputBox("Enter the sheet name (e.g., edited):", "Sheet Name", "perpunuar.")
    
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
    
    ' Loop through each row in column H of the sheet
    For i = 2 To lastRow ' Start from row 2 to skip headers
        cellValue = Trim(ws.Cells(i, 8).Value) ' Column H - Trim spaces
        
        ' Remove numbers and question marks
        processedText = cellValue
        processedText = Application.WorksheetFunction.Substitute(processedText, "0", "")
        processedText = Application.WorksheetFunction.Substitute(processedText, "1", "")
        processedText = Application.WorksheetFunction.Substitute(processedText, "2", "")
        processedText = Application.WorksheetFunction.Substitute(processedText, "3", "")
        processedText = Application.WorksheetFunction.Substitute(processedText, "4", "")
        processedText = Application.WorksheetFunction.Substitute(processedText, "5", "")
        processedText = Application.WorksheetFunction.Substitute(processedText, "6", "")
        processedText = Application.WorksheetFunction.Substitute(processedText, "7", "")
        processedText = Application.WorksheetFunction.Substitute(processedText, "8", "")
        processedText = Application.WorksheetFunction.Substitute(processedText, "9", "")
        processedText = Application.WorksheetFunction.Substitute(processedText, "?", "")
        processedText = Application.WorksheetFunction.Substitute(processedText, ".", "")
        processedText = Application.WorksheetFunction.Substitute(processedText, "_", "")
        processedText = Application.WorksheetFunction.Substitute(processedText, "-", "")
        processedText = Application.WorksheetFunction.Substitute(processedText, "(", "")
        processedText = Application.WorksheetFunction.Substitute(processedText, ")", "")
        processedText = Application.WorksheetFunction.Substitute(processedText, ",", "")
        processedText = Application.WorksheetFunction.Substitute(processedText, "!", "")
        processedText = Application.WorksheetFunction.Substitute(processedText, "@", "")
        processedText = Application.WorksheetFunction.Substitute(processedText, "'", "")
        
        ' Update column H with processed text
        ws.Cells(i, 8).Value = processedText
    Next i
    
    MsgBox "Processing complete. Numbers and question marks removed from column H."
End Sub
