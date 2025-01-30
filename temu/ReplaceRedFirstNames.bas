Sub ReplaceRedFirstNames()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim splitParts As Variant
    Dim randomName As String
    Dim lastName As String
    Dim randomNames As Variant
    Dim colorCode As Long
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Find the last row in column F
    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row ' Column F is the 6th column
    
    ' Define red background color used for highlighting (RGB 255, 0, 0)
    colorCode = RGB(255, 0, 0)
    
    ' Define a list of random Albanian first names
    randomNames = Array("Ardit", "Blerim", "Dritan", "Elton", "Florian", "Gerti", "Ilir", "Jetmir", "Isa", "Edi", _
                        "Marin", "Nertil", "Orgest", "Anisa", "Fjori", "Sokol", "Tedi", "Uran", "Valon", "Juliana", _
                        "Edison", "Zamir", "Adrian", "Besnik", "Endrit", "Fatjon", "Genc", "Ismail", "Julian", _
                        "Besnik", "Ledion", "Nertil", "Nikolin", "Orest", "Petrit", "Qamil", "Roland", "Shkëlzen", _
                        "Marsi", "Alda", "Besa",  "Alba","Ana","Elda","Altin", "Bardhyl", "Ada", "Edison", "Fation", "Gezim","Erjon","Erjola","Anxhela")

    ' Loop through column F starting from row 19
    For Each cell In ws.Range("F19:F" & lastRow)
        ' Check if the row is highlighted in red
        If cell.Interior.Color = colorCode Then
            ' Split the full name into first name and last name
            splitParts = Split(Trim(cell.Value), " ")
            
            ' Ensure the name has at least two parts
            If UBound(splitParts) >= 1 Then
                ' Keep the last word as the surname
                lastName = splitParts(UBound(splitParts))
                
                ' Select a random first name (duplicates allowed)
                randomName = randomNames(Int((UBound(randomNames) + 1) * Rnd))
                
                ' Replace the name in the cell
                cell.Value = randomName & " " & lastName
            End If
        End If
    Next cell
    
    MsgBox "Highlighted names have been updated with new first names while keeping the surname.", vbInformation
End Sub
