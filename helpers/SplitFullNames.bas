Sub SplitFullNames()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim splitParts As Variant
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Find the last row in column F
    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row ' Column F is the 6th column
    
    ' Loop through column F from row 1 and below
    For Each cell In ws.Range("F1:F" & lastRow)
        If Trim(cell.Value) <> "" Then
            ' Split the full name by space
            splitParts = Split(cell.Value, " ")
            
            ' Store first name in column G
            ws.Cells(cell.Row, 7).Value = splitParts(0)
            
            ' Store last name in column H
            If UBound(splitParts) >= 1 Then
                ws.Cells(cell.Row, 8).Value = splitParts(UBound(splitParts))
            Else
                ws.Cells(cell.Row, 8).Value = "" ' If no surname exists, leave it empty
            End If
        End If
    Next cell
    
    MsgBox "Full names split into columns G (First Name) and H (Surname).", vbInformation
End Sub

