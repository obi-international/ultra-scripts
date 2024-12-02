Sub ProcessSingleWordNamesAndAddSurnames()
    Dim sourceWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim surnames As Variant
    Dim randomSurname As String
    Dim processedText As String
    
    ' Set the source worksheet
    Set sourceWs = ThisWorkbook.Sheets("edited-12-2-2024")
    
    ' Find the last row in column F
    lastRow = sourceWs.Cells(sourceWs.Rows.Count, 6).End(xlUp).Row ' Column F is the 6th column
    
    ' Define the list of random surnames
    surnames = Array("Hoxha", "Koli", "Dajti", "Berisha", "Shala", "Gashi", "Leka", "Rama", "Meta", _
                     "Kryemadhi", "Çaka", "Bajrami", "Meksi", "Vokshi", "Muça", "Bushi", "Gjika", "Nushi", _
                     "Rrumbullaku", "Pasha")
    
    ' Loop through each row in column F of the source sheet
    For i = 2 To lastRow ' Start from row 2 to skip headers
        cellValue = Trim(sourceWs.Cells(i, 6).Value) ' Column F - Trim spaces
        
        ' Check if the cell contains exactly one word (no spaces)
        If UBound(Split(cellValue, " ")) = 0 Then
            ' Assign a random surname if the entry is only one word
            randomSurname = surnames(Int((UBound(surnames) + 1) * Rnd))
            processedText = cellValue & " " & randomSurname
            
            ' Update column F with the new name (with surname)
            sourceWs.Cells(i, 6).Value = processedText
        End If
    Next i
End Sub