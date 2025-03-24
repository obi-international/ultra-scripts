Sub ProcessUnderScorePointAndCreateNewSheet()
    Dim sourceWs As Worksheet, targetWs As Worksheet
    Dim lastRow As Long, targetRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim processedText As String
    Dim sourceSheetName As String
    Dim targetSheetName As String
    
    ' Prompt the user to enter the source and target sheet names
    sourceSheetName = InputBox("Enter the source sheet name (e.g., edited):", "Source Sheet Name", "edited")
    targetSheetName = InputBox("Enter the target sheet name (e.g., edited.):", "Target Sheet Name", "edited.")
    
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
    
    ' Create or clear the target worksheet
    On Error Resume Next
    Set targetWs = ThisWorkbook.Sheets(targetSheetName)
    If targetWs Is Nothing Then
        Set targetWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        targetWs.Name = targetSheetName
    Else
        targetWs.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Copy headers if applicable (e.g., first row)
    sourceWs.Rows(1).Copy Destination:=targetWs.Rows(1)
    
    ' Initialize target row
    targetRow = 2 ' Assuming the first row is headers
    
    ' Loop through each row in column H of the source sheet
    For i = 2 To lastRow ' Start from row 2 to skip headers
        cellValue = Trim(sourceWs.Cells(i, 8).Value) ' Column H - Trim spaces
        
        ' Replace underscores `_` and points `.` with spaces
        cellValue = Replace(cellValue, "_", " ")
        cellValue = Replace(cellValue, ".", " ")
        
        ' Convert the processed text to Proper Case
        processedText = Application.WorksheetFunction.Proper(LCase(cellValue))
        
        ' Copy the entire row to the new sheet
        sourceWs.Rows(i).Copy Destination:=targetWs.Rows(targetRow)
        
        ' Update column H with processed text
        targetWs.Cells(targetRow, 8).Value = processedText
        
        ' Increment target row counter
        targetRow = targetRow + 1
    Next i
End Sub
