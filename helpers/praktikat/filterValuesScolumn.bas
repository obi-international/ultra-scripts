Sub FilterAndCopyValues()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim sourceSheetName As String, destSheetName As String, filterColumn As String
    Dim lastRow As Long, i As Long, destRow As Long
    Dim filterColNum As Long
    
    ' Set default values
    sourceSheetName = "initially" ' Default source sheet name
    filterColumn = "O" ' Default column to filter
    destSheetName = "ready" ' Default destination sheet name
    
    ' Prompt for source sheet name (user can change it)
    sourceSheetName = InputBox("Enter the source sheet name:", "Source Sheet", sourceSheetName)
    
    ' Check if the source sheet exists
    On Error Resume Next
    Set wsSource = ThisWorkbook.Sheets(sourceSheetName)
    On Error GoTo 0
    
    If wsSource Is Nothing Then
        MsgBox "Source sheet not found. Please check the name and try again.", vbCritical
        Exit Sub
    End If
    
    ' Prompt for filter column letter (default S)
    filterColumn = InputBox("Enter the column letter to filter:", "Filter Column", filterColumn)
    
    ' Convert column letter to column number
    filterColNum = Columns(filterColumn).Column
    
    ' Prompt for destination sheet name (default "ready")
    destSheetName = InputBox("Enter the destination sheet name:", "Destination Sheet", destSheetName)
    
    ' Check if the destination sheet exists, create if not
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets(destSheetName)
    On Error GoTo 0
    
    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Sheets.Add
        wsDest.Name = destSheetName
        MsgBox "Destination sheet not found, so a new one was created: " & destSheetName, vbInformation
    End If
    
    ' Find the last row in the filter column
    lastRow = wsSource.Cells(wsSource.Rows.Count, filterColNum).End(xlUp).Row
    
    ' Clear destination sheet before pasting
    wsDest.Cells.Clear
    
    ' Copy headers (assuming headers are in row 1)
    wsSource.Rows(1).Copy
    wsDest.Rows(1).PasteSpecial Paste:=xlPasteValues
    destRow = 2 ' Start pasting from row 2
    
    ' Loop through each row and filter based on the specified column value = 1
    For i = 2 To lastRow ' Assuming data starts from row 2
        If wsSource.Cells(i, filterColNum).Value > 1 Then
            ' Copy the entire row
            wsSource.Rows(i).Copy
            
            ' Paste only values into the destination sheet
            wsDest.Rows(destRow).PasteSpecial Paste:=xlPasteValues
            
            destRow = destRow + 1 ' Move to the next row
        End If
    Next i
    
    ' Clear clipboard to avoid issues
    Application.CutCopyMode = False
    
    MsgBox "Filtered data copied as values to " & destSheetName, vbInformation
End Sub
