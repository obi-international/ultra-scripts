' In Module: NameArrayModule
Public Function GetRandomName() As String
    Dim names As Variant
    Dim randomIndex As Integer
    Dim lastRow As Long

    ' Get the last row with data in column M
    lastRow = Cells(Rows.Count, "M").End(xlUp).Row
    
    ' Pull all names from column M into the names array
    names = Range("M1:M" & lastRow).Value
    
    ' Generate a random index (1 to lastRow)
    Randomize ' Ensure different results on each run
    randomIndex = Int((lastRow - 1 + 1) * Rnd) + 1
    
    ' Return a single random name
    GetRandomName = names(randomIndex, 1)
End Function

Sub TestRandomName()
    MsgBox "Random Name: " & GetRandomName()
End Sub
