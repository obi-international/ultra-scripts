Sub ReplaceRedFirstNames()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim splitParts As Variant
    Dim randomName As String
    Dim lastName As String
    Dim colorCode As Long
    Dim names As Variant
    Dim randomIndex As Integer
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Find the last row in column F
    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row ' Column F is the 6th column
    
    ' Define red background color used for highlighting (RGB 255, 0, 0)
    colorCode = RGB(255, 0, 0)

    ' Get the last row with data in column M
    lastRow = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row
    
    ' Pull all names from column M into the names array
    names = ws.Range("M1:M" & lastRow).Value
    
    ' Generate a random index (1 to lastRow)
    Randomize ' Ensure different results on each run
    randomIndex = Int((lastRow - 1 + 1) * Rnd) + 1
    
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
                
                ' Select a random first name from column M
                randomName = names(randomIndex, 1)
                
                ' Replace the name in the cell with the new first name and last name
                cell.Value = randomName & " " & lastName
            End If
        End If
    Next cell
    
    MsgBox "Highlighted names have been updated with new first names while keeping the surname.", vbInformation
End Sub
