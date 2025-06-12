Sub FilterRedRows()
    Dim ws As Worksheet, newWs As Worksheet
    Dim lastRow As Long, i As Long, destRow As Long
    Dim sourceSheetName As String, targetSheetName As String
    
    ' Ask user for the source sheet name
    sourceSheetName = InputBox("Enter the sheet name to process:", "Source Sheet Selection", "NotYellow")
    
    ' Check if the source sheet exists
    On Error Resume Next
    Set ws = Sheets(sourceSheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet '" & sourceSheetName & "' not found. Please check the name and try again.", vbCritical
        Exit Sub
    End If
    
    ' Ask user for the target sheet name
    targetSheetName = InputBox("Enter the sheet name where filtered rows will be saved:", "Target Sheet Selection", "init-2+")
    
    ' Create or find the target sheet
    On Error Resume Next
    Set newWs = Sheets(targetSheetName)
    If newWs Is Nothing Then
        Set newWs = Sheets.Add
        newWs.Name = targetSheetName
    End If
    On Error GoTo 0
    
    ' Clear the target sheet before copying
    newWs.Cells.Clear
    
    ' Find the last row in column F of the source sheet
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    
    ' Copy headers from row 1
    ws.Rows(1).Copy Destination:=newWs.Rows(1)
    destRow = 2 ' Start copying from row 2 in the target sheet
    
    ' Loop through rows starting from F3 downwards
    For i = 3 To lastRow
        If ws.Cells(i, 6).Interior.Color = RGB(255, 0, 0) Then ' Check if column F is red
            ws.Rows(i).Copy Destination:=newWs.Rows(destRow)
            destRow = destRow + 1
        End If
    Next i
    
    MsgBox "Filtering complete! The rows with a red background in column F are now in '" & targetSheetName & "' sheet.", vbInformation
End Sub

