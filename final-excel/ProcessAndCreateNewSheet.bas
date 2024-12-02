Sub ProcessAndCreateNewSheet()
    Dim sourceWs As Worksheet, targetWs As Worksheet
    Dim lastRow As Long, targetRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim splitParts As Variant
    Dim processedText As String
    Dim sourceSheetName As String
    Dim targetSheetName As String
    
    ' Prompt the user to enter the source and target sheet names
    sourceSheetName = InputBox("Enter the source sheet name (e.g., 12-2-2024):", "Source Sheet Name", "12-2-2024")
    targetSheetName = InputBox("Enter the target sheet name (e.g., edited-12-2-2024):", "Target Sheet Name", "edited-12-2-2024")
    
    ' Set the source worksheet
    On Error Resume Next
    Set sourceWs = ThisWorkbook.Sheets(sourceSheetName)
    On Error GoTo 0
    If sourceWs Is Nothing Then
        MsgBox "Source sheet not found!"
        Exit Sub
    End If
    
    ' Find the last row in column F
    lastRow = sourceWs.Cells(sourceWs.Rows.Count, 6).End(xlUp).Row ' Column F is the 6th column
    
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
    
    ' Loop through each row in column F of the source sheet
    For i = 2 To lastRow ' Start from row 2 to skip headers
        cellValue = Trim(sourceWs.Cells(i, 6).Value) ' Column F - Trim spaces
        
        ' Handle slashes `/` - keep only the first part
        If InStr(cellValue, "/") > 0 Then
            cellValue = Split(cellValue, "/")(0)
        End If
        
        ' Handle periods `.` in the text - split on periods and keep only the first two words
        If InStr(cellValue, ".") > 0 Then
            splitParts = Split(cellValue, ".")
            If UBound(splitParts) >= 1 Then
                processedText = splitParts(0) & " " & splitParts(1)
            Else
                processedText = splitParts(0)
            End If
        ElseIf InStr(cellValue, "_") > 0 Then
            ' Handle underscores `_` - replace them with spaces and keep the first two words
            splitParts = Split(cellValue, "_")
            If UBound(splitParts) >= 1 Then
                processedText = splitParts(0) & " " & splitParts(1)
            Else
                processedText = splitParts(0)
            End If
        Else
            ' If no period, underscore, or slash, just trim spaces
            processedText = cellValue
        End If
        
        ' Convert the processed text to lowercase and capitalize the first letter of each word
        processedText = Application.WorksheetFunction.Proper(LCase(processedText))
        
        ' Copy the entire row to the new sheet
        sourceWs.Rows(i).Copy Destination:=targetWs.Rows(targetRow)
        
        ' Update column F with processed text
        targetWs.Cells(targetRow, 6).Value = processedText
        
        ' Increment target row counter
        targetRow = targetRow + 1
    Next i
End Sub
