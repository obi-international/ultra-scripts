Sub RemoveSpacesFromColumn()
    Dim sheetName As String
    Dim colLetter As String
    Dim colNum As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Prompt for sheet name
    sheetName = InputBox("Enter the sheet name:", "Sheet Name","HS Codes")
    If sheetName = "" Then Exit Sub

    ' Prompt for column letter
    colLetter = InputBox("Enter the column letter to remove spaces from:", "Column Letter","I")
    If colLetter = "" Then Exit Sub

    ' Convert column letter to number
    On Error Resume Next
    colNum = Range(colLetter & "1").Column
    On Error GoTo 0
    If colNum = 0 Then
        MsgBox "Invalid column letter!", vbExclamation
        Exit Sub
    End If

    ' Get the worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' not found!", vbCritical
        Exit Sub
    End If

    ' Find the last row
    lastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row

    ' Loop and remove spaces
    For i = 2 To lastRow ' Assuming row 1 is headers
        If Not IsEmpty(ws.Cells(i, colNum).Value) Then
            ws.Cells(i, colNum).Value = Replace(ws.Cells(i, colNum).Value, " ", "")
        End If
    Next i

    MsgBox "Spaces removed from column " & colLetter & " in sheet '" & sheetName & "'.", vbInformation
End Sub

