Sub ProcessAndCreateNewSheet()
    Dim sourceWs As Worksheet, targetWs As Worksheet
    Dim lastRow As Long, targetRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim processedText As String
    Dim sourceSheetName As String
    Dim targetSheetName As String
    Dim parts As Variant
    
    ' Prompt the user to enter the source and target sheet names
    sourceSheetName = InputBox("Enter the source sheet name (e.g., 12-2-2024):", "Source Sheet Name", "original")
    targetSheetName = InputBox("Enter the target sheet name (e.g., edited-12-2-2024):", "Target Sheet Name", "edited")
    
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
        
        ' Handle slashes `/` or dashes `-`
        If InStr(cellValue, "/") > 0 Or InStr(cellValue, "-") > 0 Then
            parts = Split(cellValue, "/") ' First split by slash
            If UBound(parts) = 0 Then parts = Split(cellValue, "-") ' If no slash, split by dash
            processedText = GetLongestPart(parts)
        Else
            processedText = cellValue
        End If
        
        ' Convert the processed text to Proper Case
        processedText = Application.WorksheetFunction.Proper(LCase(processedText))
        
        ' Copy the entire row to the new sheet
        sourceWs.Rows(i).Copy Destination:=targetWs.Rows(targetRow)
        
        ' Update column H with processed text
        targetWs.Cells(targetRow, 8).Value = processedText
        
        ' Increment target row counter
        targetRow = targetRow + 1
    Next i
End Sub

' Function to get the longest part of the split text
Function GetLongestPart(parts As Variant) As String
    Dim i As Long
    Dim longest As String
    Dim maxLength As Long
    
    For i = LBound(parts) To UBound(parts)
        If Len(Trim(parts(i))) > maxLength Then
            longest = Trim(parts(i))
            maxLength = Len(Trim(parts(i)))
        End If
    Next i
    
    GetLongestPart = longest
End Function
