Sub TranslateColumn()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim sheetName As String
    
    ' Prompt the user to input the sheet name
    sheetName = InputBox("Enter the name of the sheet where translation is needed:", "Sheet Name")
    
    ' Check if the sheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "The sheet '" & sheetName & "' does not exist. Please check and try again.", vbCritical
        Exit Sub
    End If
    
    ' Find the last row with data in column I (English descriptions)
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    
    ' Loop through the rows to translate and fill column H
    For i = 19 To lastRow
        If ws.Cells(i, "I").Value <> "" Then
            ws.Cells(i, "H").Value = TranslateText(ws.Cells(i, "I").Value, "en", "it")
        End If
    Next i
    
    MsgBox "Translation completed for sheet: " & sheetName, vbInformation
End Sub
