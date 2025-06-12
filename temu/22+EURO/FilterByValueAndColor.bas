Sub FilterByValueAndColor()
    Dim ws As Worksheet, newWs As Worksheet
    Dim lastRow As Long, i As Long, destRow As Long
    Dim sourceSheetName As String, targetSheetName As String
    Dim filterColumnLetter As String
    Dim filterColNum As Long
    Dim cellValue As Variant
    Dim cellColor As Long

    ' Prompt for source sheet
    sourceSheetName = InputBox("Enter the source sheet name:", "Source Sheet", "e")
    If sourceSheetName = "" Then Exit Sub
    
    On Error Resume Next
    Set ws = Sheets(sourceSheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet '" & sourceSheetName & "' not found.", vbCritical
        Exit Sub
    End If

    ' Prompt for target sheet
    targetSheetName = InputBox("Enter the target sheet name:", "Target Sheet", "filtered")
    
    On Error Resume Next
    Set newWs = Sheets(targetSheetName)
    If newWs Is Nothing Then
        Set newWs = Sheets.Add
        newWs.Name = targetSheetName
    End If
    On Error GoTo 0
    
    ' Prompt for column (default "O")
    filterColumnLetter = InputBox("Enter the column letter to filter by (e.g., O):", "Column Letter", "O")
    If filterColumnLetter = "" Then Exit Sub
    filterColNum = Columns(filterColumnLetter).Column

    ' Clear target sheet
    newWs.Cells.Clear

    ' Copy header
    ws.Rows(1).Copy Destination:=newWs.Rows(1)
    destRow = 2

    ' Get last row
    lastRow = ws.Cells(ws.Rows.Count, filterColNum).End(xlUp).Row

    ' Loop through each row starting from row 2
    For i = 2 To lastRow
        cellValue = ws.Cells(i, filterColNum).Value
        cellColor = ws.Cells(i, filterColNum).Interior.Color

        If IsNumeric(cellValue) Then
            If cellValue >= 22 And cellColor <> RGB(255, 0, 0) Then
                ws.Rows(i).Copy Destination:=newWs.Rows(destRow)
                destRow = destRow + 1
            End If
        End If
    Next i

    MsgBox "Filtered rows copied to '" & targetSheetName & "'.", vbInformation
End Sub
