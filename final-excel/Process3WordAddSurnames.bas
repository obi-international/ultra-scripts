Sub ProcessWordAddSurnames()
    Dim sourceWs As Worksheet, targetWs As Worksheet
    Dim lastRow As Long, targetRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim surnames As Variant
    Dim randomSurname As String
    Dim processedText As String
    Dim splitParts As Variant
    Dim sourceSheetName As String
    Dim targetSheetName As String
    
    ' Prompt the user to enter the source sheet name
    sourceSheetName = InputBox("Enter the source sheet name (e.g., edited.):", "Source Sheet Name", "edited.")
    targetSheetName = InputBox("Enter the target sheet name (default: perpunuar):", "Target Sheet Name", "perpunuar")
    
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
    
    ' Define the list of random surnames
    surnames = Array("Hoxha", "Paja", "Dajti", "Balliu", "Shala", "Hoxhaj", "Leka", "Kaci", "Muca")
    
    ' Loop through each row in column H of the source sheet
    For i = 2 To lastRow ' Start from row 2 to skip headers
        cellValue = Trim(sourceWs.Cells(i, 8).Value) ' Column H - Trim spaces
        
        ' Split the cell value by space to check for the number of words
        splitParts = Split(cellValue, " ")
        
        ' Case 1: If there is exactly one word (no spaces)
        If UBound(splitParts) = 0 Then
            ' Assign a random surname if the entry is only one word
            randomSurname = surnames(Int((UBound(surnames) + 1) * Rnd))
            processedText = cellValue & " " & randomSurname
        
        ' Case 2: If there are more than two words, keep only the first two
        ElseIf UBound(splitParts) >= 2 Then
            processedText = splitParts(0) & " " & splitParts(1)
        
        ' Case 3: If the entry already has two words, keep it as is
        Else
            processedText = cellValue
        End If
        
        ' Write the processed text to the target sheet
        sourceWs.Rows(i).Copy Destination:=targetWs.Rows(targetRow)
        targetWs.Cells(targetRow, 8).Value = processedText
        
        ' Increment target row counter
        targetRow = targetRow + 1
    Next i
End Sub
