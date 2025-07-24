Sub CountItemsByPipeDelimiter()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim descColLetter As String
    Dim outColLetter As String
    Dim lastRow As Long
    Dim i As Long
    Dim descText As String
    Dim parts As Variant

    ' Prompt for inputs
    sheetName = InputBox("Enter the sheet name to process:", "Sheet Name", "teke")
    If sheetName = "" Then Exit Sub

    descColLetter = InputBox("Enter the column letter for description (default = F):", "Description Column", "F")
    If descColLetter = "" Then descColLetter = "F"

    outColLetter = InputBox("Enter the column letter for output count (default = N):", "Output Column", "K")
    If outColLetter = "" Then outColLetter = "N"

    ' Set worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' not found!", vbCritical
        Exit Sub
    End If

    ' Get last row in description column
    lastRow = ws.Cells(ws.Rows.Count, descColLetter).End(xlUp).Row

    ' Loop through each row
    For i = 2 To lastRow
        descText = Trim(ws.Cells(i, descColLetter).Value)
        If descText <> "" Then
            parts = Split(descText, "|")
            ws.Cells(i, outColLetter).Value = UBound(parts) - LBound(parts) + 1
        Else
            ws.Cells(i, outColLetter).Value = 0
        End If
    Next i

    MsgBox "Item counts written to column " & outColLetter & ".", vbInformation
End Sub
