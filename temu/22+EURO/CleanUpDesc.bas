Sub CleanUpDescriptions()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim colLetter As String
    Dim lastRow As Long, i As Long
    Dim fullText As String, parts As Variant

    sheetName = InputBox("Enter the sheet name:", "Sheet Name", "teke")
    If sheetName = "" Then Exit Sub
    
    colLetter = InputBox("Enter the column letter to clean (e.g., E):", "Column", "E")
    If colLetter = "" Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet not found!", vbExclamation
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, colLetter).End(xlUp).Row
    
    For i = 2 To lastRow
        fullText = Trim(ws.Cells(i, colLetter).Value)
        
        If fullText <> "" Then
            If InStr(fullText, " | ") > 0 Then
                parts = Split(fullText, " | ")
                ws.Cells(i, colLetter).Value = Trim(parts(0))
            Else
                ' No delimiter, keep the value as is
                ws.Cells(i, colLetter).Value = fullText
            End If
        End If
    Next i

    MsgBox "Column cleaned. First item kept, or original value retained if no delimiter found.", vbInformation
End Sub
