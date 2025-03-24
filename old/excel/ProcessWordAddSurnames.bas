Sub ProcessWordAddSurnames()
    Dim sourceWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim surnames As Variant
    Dim randomSurname As String
    Dim processedText As String
    Dim splitParts As Variant
    Dim sourceSheetName As String
    
    ' Prompt the user to enter the source sheet name
    sourceSheetName = InputBox("Enter the source sheet name (e.g., edited-12-2-2024):", "Source Sheet Name", "edited-12-2-2024")
    
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
    
    ' Define the list of random surnames
    surnames = Array("Hoxha", "Paja", "Dajti", "Balliu", "Shala", "Hoxhaj", "Leka", "Kaci", "Muca")
    
    ' Loop through each row in column F of the source sheet
    For i = 2 To lastRow ' Start from row 2 to skip headers
        cellValue = Trim(sourceWs.Cells(i, 6).Value) ' Column F - Trim spaces
        
        ' Split the cell value by space to check for the number of words
        splitParts = Split(cellValue, " ")
        
        ' Case 1: If there is exactly one word (no spaces)
        If UBound(splitParts) = 0 Then
            ' Assign a random surname if the entry is only one word
            randomSurname = surnames(Int((UBound(surnames) + 1) * Rnd))
            processedText = cellValue & " " & randomSurname
            sourceWs.Cells(i, 6).Value = processedText
        
        ' Case 2: If there are more than two words, keep only the first two
        ElseIf UBound(splitParts) >= 2 Then
            processedText = splitParts(0) & " " & splitParts(1)
            sourceWs.Cells(i, 6).Value = processedText
        End If
    Next i
End Sub