Sub CopyNonRedRowsToNewSheet()
    Dim sourceSheetName As String
    sourceSheetName = InputBox("Enter the sheet name to copy from:", "Sheet name", "NotYellow")
    If sourceSheetName = "" Then Exit Sub

    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim lastRow As Long, i As Long, targetRow As Long
    Dim redColor As Long

    redColor = vbRed ' Standard red color constant

    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets(sourceSheetName)
    On Error GoTo 0

    If wsSource Is Nothing Then
        MsgBox "Sheet '" & sourceSheetName & "' not found!", vbExclamation
        Exit Sub
    End If

    ' Delete "NotRed" if it exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("NotRed").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Create new target sheet
    Set wsTarget = ThisWorkbook.Worksheets.Add
    wsTarget.Name = "NotRed"

    ' Copy headers
    wsSource.Rows(1).Copy Destination:=wsTarget.Rows(1)
    targetRow = 2

    lastRow = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row

    For i = 2 To lastRow
        If wsSource.Cells(i, "D").Interior.Color <> redColor Then
            wsSource.Rows(i).Copy Destination:=wsTarget.Rows(targetRow)
            targetRow = targetRow + 1
        End If
    Next i

    MsgBox "Rows without red cells in column D copied to 'NotRed' sheet.", vbInformation
End Sub
