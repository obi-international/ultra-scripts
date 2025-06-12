Sub CopyYellowRowsToNewSheet()
    Dim sourceSheetName As String
    sourceSheetName = InputBox("Enter the sheet name to copy from:", "Sheet name", "filtered")
    If sourceSheetName = "" Then Exit Sub

    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim lastRow As Long, i As Long, targetRow As Long
    Dim cellColor As Long

    cellColor = vbYellow ' Standard yellow color constant

    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets(sourceSheetName)
    On Error GoTo 0

    If wsSource Is Nothing Then
        MsgBox "Sheet '" & sourceSheetName & "' not found!", vbExclamation
        Exit Sub
    End If

    ' Delete "1 item" if it exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("1 item").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Create new target sheet
    Set wsTarget = ThisWorkbook.Worksheets.Add
    wsTarget.Name = "1 item"

    ' Copy headers
    wsSource.Rows(1).Copy Destination:=wsTarget.Rows(1)
    targetRow = 2

    lastRow = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row

    For i = 2 To lastRow
        If wsSource.Cells(i, "D").Interior.Color = cellColor Then
            wsSource.Rows(i).Copy Destination:=wsTarget.Rows(targetRow)
            targetRow = targetRow + 1
        End If
    Next i

    MsgBox "Rows with yellow cells in column D copied to '1 item' sheet.", vbInformation
End Sub