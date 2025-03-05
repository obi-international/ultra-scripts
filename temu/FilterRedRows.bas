Sub FilterRedRows()
' todo add the request the sheet name to add those
    Dim ws As Worksheet, newWs As Worksheet
    Dim lastRow As Long, i As Long, destRow As Long
    Dim sheetName As String
    
    ' Ask user for the sheet name
    sheetName = InputBox("Enter the sheet name to process:", "Sheet Selection","e")
    
    ' Check if the sheet name exists
    On Error Resume Next
    Set ws = Sheets(sheetName)
    On Error GoTo 0
    
    ' If the sheet does not exist, show an error message and exit
    If ws Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' not found. Please check the name and try again.", vbCritical
        Exit Sub
    End If
    
    ' Find the last row in column F
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    
    ' Create or find the 'initially' sheet
    On Error Resume Next
    Set newWs = Sheets("initially")
    If newWs Is Nothing Then
        Set newWs = Sheets.Add
        newWs.Name = "initially"
    End If
    On Error GoTo 0
    
    ' Clear the new sheet before copying
    newWs.Cells.Clear
    
    ' Copy headers from row 1
    ws.Rows(1).Copy Destination:=newWs.Rows(1)
    destRow = 2  ' Start copying from row 2 in the 'initially' sheet
    
    ' Loop through rows starting from F3 downwards
    For i = 3 To lastRow
        If ws.Cells(i, 6).Interior.Color = RGB(255, 0, 0) Then ' Check if F column is red
            ws.Rows(i).Copy Destination:=newWs.Rows(destRow)
            destRow = destRow + 1
        End If
    Next i
    
    MsgBox "Filtering complete! The rows with a red background in column F are now in 'initially' sheet.", vbInformation
End Sub
