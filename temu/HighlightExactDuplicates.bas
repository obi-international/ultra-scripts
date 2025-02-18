Sub HighlightExactDuplicates()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim nameCounts As Object
    Dim cellValue As String
    Dim i As Long
    Dim sheetName As String
    
    ' Prompt the user to enter the sheet name
    sheetName = InputBox("Enter the sheet name (e.g., original):", "Sheet Name", "e")
    
    ' Set the worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet not found!", vbExclamation
        Exit Sub
    End If
    
    ' Find the last row in column F
    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row ' Column F is the 6th column
    
    ' Create a dictionary to count names
    Set nameCounts = CreateObject("Scripting.Dictionary")
    
    ' Loop through column F starting from row 3 to count names
    For i = 3 To lastRow ' Start from row 3
        cellValue = Trim(ws.Cells(i, 6).Value) ' Column F - Trim spaces
        If cellValue <> "" Then
            If nameCounts.exists(cellValue) Then
                nameCounts(cellValue) = nameCounts(cellValue) + 1
            Else
                nameCounts.Add cellValue, 1
            End If
        End If
    Next i
    
    ' Loop through column F again to highlight rows where names appear exactly 2 times
    For i = 3 To lastRow ' Start from row 3
        cellValue = Trim(ws.Cells(i, 6).Value) ' Column F - Trim spaces
        If cellValue <> "" Then
            If nameCounts(cellValue) = 2 Then ' Only highlight if count is exactly 2
                ws.Rows(i).Interior.Color = RGB(255, 0, 0) ' Highlight row in red
            End If
        End If
    Next i
    
    MsgBox "Rows with names appearing exactly 2 times have been highlighted in red.", vbInformation
End Sub
