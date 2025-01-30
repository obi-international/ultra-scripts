Sub ProcessNameAndSurnameStartingAtF19()
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
    sheetName = InputBox("Enter the sheet name (e.g., original):", "Sheet Name", "23-12-2024-1")
    
    ' Set the worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet not found!"
        Exit Sub
    End If
    
    ' Find the last row in column F
    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row ' Column F is the 6th column
    
    ' Define the list of random surnames
    surnames = Array("Hoxha", "Paja", "Dajti", "Balliu", "Shala", "Hoxhaj", "Leka", "Kaci", "Muca", "Hysa", "Tafani", "Qosja", "Ismaili", "Bali", "Marku", "Tusha", "Bajrami", "Domi", "Doshi", "Berberi", "Xhafa","Gjoci","Kapaj","Laska","Meta","Haka","Kasa","Gjoka","Gjoni")
    
    ' Loop through each row in column F starting from F19
    For i = 19 To lastRow ' Start from row 19
        cellValue = Trim(ws.Cells(i, 6).Value) ' Column F - Trim spaces
        
        ' Split the cell value by space to check for the number of words
        splitParts = Split(cellValue, " ")
        
        ' Case 1: If there is exactly one word (no spaces)
        If UBound(splitParts) = 0 Then
            ' Assign a random surname if the entry is only one word
            randomSurname = surnames(Int((UBound(surnames) + 1) * Rnd))
            processedText = Application.WorksheetFunction.Proper(LCase(splitParts(0))) & " " & randomSurname
        
        ' Case 2: If there are more than two words, keep only the first two
        ElseIf UBound(splitParts) >= 2 Then
            processedText = Application.WorksheetFunction.Proper(LCase(splitParts(0))) & " " & Application.WorksheetFunction.Proper(LCase(splitParts(1)))
        
        ' Case 3: If the entry already has two words, format them properly
        Else
            processedText = Application.WorksheetFunction.Proper(LCase(splitParts(0))) & " " & Application.WorksheetFunction.Proper(LCase(splitParts(1)))
        End If
        
        ' Update the cell in the same column with the processed text
        ws.Cells(i, 6).Value = processedText
    Next i
    
    MsgBox "Processing complete. Column F starting at F19 has been updated with proper names and surnames.", vbInformation
End Sub
