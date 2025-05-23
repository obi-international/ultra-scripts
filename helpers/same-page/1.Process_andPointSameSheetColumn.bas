Sub ProcessUnderScorePointInSameSheet()
    Dim sourceWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim processedText As String
    Dim sourceSheetName As String
    
    ' Prompt the user to enter the source sheet name
    sourceSheetName = InputBox("Enter the source sheet name (e.g., edited):", "Source Sheet Name", "original")
    
    ' Set the source worksheet
    On Error Resume Next
    Set sourceWs = ThisWorkbook.Sheets(sourceSheetName)
    On Error GoTo 0
    If sourceWs Is Nothing Then
        MsgBox "Source sheet not found!"
        Exit Sub
    End If
    
    ' Find the last row in column H
    lastRow = sourceWs.Cells(sourceWs.Rows.Count, 8).End(xlUp).Row ' Column H is the 8th column
    
    ' Loop through each row in column H of the source sheet
    For i = 2 To lastRow ' Start from row 2 to skip headers
        cellValue = Trim(sourceWs.Cells(i, 8).Value) ' Column H - Trim spaces
        
        ' Replace underscores `_` and points `.` with spaces
        cellValue = Replace(cellValue, "_", " ")
        cellValue = Replace(cellValue, ".", " ")
        
        ' Convert the processed text to Proper Case
        processedText = Application.WorksheetFunction.Proper(LCase(cellValue))
        
        ' Update column H with the processed text
        sourceWs.Cells(i, 8).Value = processedText
    Next i
    
    MsgBox "Processing complete. Column H has been updated.", vbInformation
End Sub
