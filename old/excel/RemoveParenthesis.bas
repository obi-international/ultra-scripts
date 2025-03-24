Sub RemoveTextInParenthesesInSameSheet()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim processedText As String
    Dim sheetName As String
    Dim openParen As Long, closeParen As Long
    
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
    
    ' Find the last row in column H
    lastRow = ws.Cells(ws.Rows.Count, 8).End(xlUp).Row ' Column H is the 8th column
    
    ' Loop through each row in column H of the sheet
    For i = 2 To lastRow ' Start from row 2 to skip headers
        cellValue = Trim(ws.Cells(i, 8).Value) ' Column H - Trim spaces
        
        ' Find and remove text inside parentheses
        Do
            openParen = InStr(cellValue, "(")
            closeParen = InStr(cellValue, ")")
            
            If openParen > 0 And closeParen > openParen Then
                cellValue = Left(cellValue, openParen - 1) & Mid(cellValue, closeParen + 1)
            Else
                Exit Do
            End If
        Loop
        
        ' Trim spaces after removing text
        processedText = Trim(cellValue)
        
        ' Update column H with processed text
        ws.Cells(i, 8).Value = processedText
    Next i
    
    MsgBox "Processing complete. Text inside parentheses has been removed from column H."
End Sub
