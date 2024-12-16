Sub ProcessWordAddSurnamesInSameSheet()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim surnames As Variant
    Dim randomSurname As String
    Dim processedText As String
    Dim splitParts As Variant
    Dim sheetName As String
    
    ' Prompt the user to enter the sheet name
    sheetName = InputBox("Enter the sheet name to edit (e.g., edited):", "Sheet Name", "edited")
    
    ' Set the worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet not found!"
        Exit Sub
    End If
    
    ' Find the last row in column H
    lastRow = ws.Cells(ws.Rows.Count, 8).End(xlUp).Row ' Column H is the 8th column
    
    ' Define the list of random surnames
    surnames = Array("Hoxha", "Paja", "Dajti", "Balliu", "Shala", "Hoxhaj", "Leka", "Kaci", "Muca")
    
    ' Loop through each row in column H
    For i = 2 To lastRow ' Start from row 2 to skip headers
        cellValue = Trim(ws.Cells(i, 8).Value) ' Column H - Trim spaces
        
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
        
        ' Write the processed text to column I (9th column)
        ws.Cells(i, 9).Value = processedText
    Next i
    
    MsgBox "Processing complete. Modified values have been added to column I."
End Sub