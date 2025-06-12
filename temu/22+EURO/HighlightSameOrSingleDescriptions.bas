Sub HighlightSameOrSingleDescriptions()
    Dim sheetName As String
    Dim descColLetter As String
    Dim descCol As Long

    ' Prompt for sheet and column
    sheetName = InputBox("Enter the sheet name:", "Sheet Name", "filtered")
    If sheetName = "" Then Exit Sub
    
    descColLetter = InputBox("Enter the column letter for descriptions (default = J):", "Description Column", "J")
    If descColLetter = "" Then descColLetter = "J"
    
    ' Convert column letter to number
    descCol = Columns(descColLetter).Column

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' not found!", vbExclamation
        Exit Sub
    End If
    
    Dim lastRow As Long, i As Long, desc As String, parts As Variant
    Dim firstValue As String, allSame As Boolean
    Dim j As Long
    
    Application.ScreenUpdating = False
    lastRow = ws.Cells(ws.Rows.Count, descCol).End(xlUp).Row
    
    For i = 2 To lastRow
        desc = Trim(ws.Cells(i, descCol).value)
        If Len(desc) > 0 Then
            parts = Split(desc, "|")
            
            ' Clean and check all values
            firstValue = Trim(parts(0))
            allSame = True
            
            For j = LBound(parts) To UBound(parts)
                If Trim(parts(j)) <> firstValue Then
                    allSame = False
                    Exit For
                End If
            Next j
            
            If allSame Then
                ws.Rows(i).Interior.Color = vbYellow
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "Done checking and highlighting.", vbInformation
End Sub

