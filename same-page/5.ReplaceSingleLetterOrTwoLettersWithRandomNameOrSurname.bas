Sub ReplaceSingleLetterOrTwoLettersWithRandomNameOrSurname()
    Dim sourceWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim splitParts As Variant
    Dim surnames As Variant, names As Variant
    Dim randomName As String, randomSurname As String
    Dim sourceSheetName As String

    ' Prompt the user to enter the source sheet name
    sourceSheetName = InputBox("Enter the source sheet name (e.g., original):", "Source Sheet Name", "original")
    
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
    
    ' Define the list of random surnames and names
    surnames = Array("Hoxha", "Paja", "Dajti", "Balliu", "Shala", "Hoxhaj", "Leka", "Kaci", "Muca")
    names = Array("Arti", "Blerina", "Ermal", "Flora", "Gentian", "Ilir", "Jonida", "Kristi", "Luan", "Mira")
    
    ' Loop through each row in column H of the source sheet
    For i = 2 To lastRow ' Start from row 2 to skip headers
        cellValue = Trim(sourceWs.Cells(i, 8).Value) ' Column H - Trim spaces
        
        ' Split the cell value by space
        splitParts = Split(cellValue, " ")
        
        ' Loop through parts of the cell value
        Dim processedParts As String
        processedParts = ""
        Dim part As Variant
        For Each part In splitParts
            If Len(part) = 1 Or Len(part) = 2 Then
                ' Replace single or two-letter parts with a random name or surname
                If processedParts = "" Then
                    ' If it's the first part, replace with a name
                    randomName = names(Int((UBound(names) + 1) * Rnd))
                    processedParts = processedParts & " " & randomName
                Else
                    ' Otherwise, replace with a surname
                    randomSurname = surnames(Int((UBound(surnames) + 1) * Rnd))
                    processedParts = processedParts & " " & randomSurname
                End If
            Else
                ' Keep the part as is if it's not a single or two-letter part
                processedParts = processedParts & " " & part
            End If
        Next part
        
        ' Remove leading/trailing spaces from the processed text
        processedParts = Trim(processedParts)
        
        ' Update the processed text in the same cell (Column H)
        sourceWs.Cells(i, 8).Value = processedParts
    Next i

    MsgBox "Processing complete. Column H has been updated.", vbInformation
End Sub
