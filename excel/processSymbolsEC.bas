Sub ReplaceSpecialCharactersInCities()
    Dim sourceWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim sourceSheetName As String
    
    ' Prompt the user to enter the source sheet name
    sourceSheetName = InputBox("Enter the source sheet name (e.g., edited.):", "Source Sheet Name", "edited.")
    
    ' Set the source worksheet
    On Error Resume Next
    Set sourceWs = ThisWorkbook.Sheets(sourceSheetName)
    On Error GoTo 0
    If sourceWs Is Nothing Then
        MsgBox "Source sheet not found!"
        Exit Sub
    End If
    
    ' Find the last row in column I
    lastRow = sourceWs.Cells(sourceWs.Rows.Count, 9).End(xlUp).Row ' Column I is the 9th column
    
    ' Loop through each row in column I of the source sheet
    For i = 2 To lastRow ' Start from row 2 to skip headers
        cellValue = sourceWs.Cells(i, 9).Value ' Get the value from column I
        
        ' Replace special characters
        cellValue = Replace(cellValue, "Ã«", "e") ' Replace Ã« with e
        cellValue = Replace(cellValue, "Ã§", "c") ' Replace Ã§ with c
        
        ' Write the cleaned value back to column I
        sourceWs.Cells(i, 9).Value = cellValue
    Next i
    
    MsgBox "Special characters replaced in column I.", vbInformation
End Sub
