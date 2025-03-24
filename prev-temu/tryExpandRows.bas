Sub ExpandRows()
    Dim ws As Worksheet
    Dim lastRow As Long, newRow As Long
    Dim i As Long, j As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    ' Loop through each row and duplicate based on "COPE" column
    newRow = 2
    For i = 2 To lastRow
        Dim name As String
        Dim piece As Integer
        Dim vlera As Double
        Dim detyrim As Double
        
        ' Ensure values are correctly retrieved and converted
        name = ws.Cells(i, 2).Value
        If IsNumeric(ws.Cells(i, 3).Value) Then
            piece = CInt(ws.Cells(i, 3).Value) ' Convert to integer
        Else
            piece = 1 ' Default to 1 if not a valid number
        End If
        If IsNumeric(ws.Cells(i, 4).Value) Then
            vlera = CDbl(ws.Cells(i, 4).Value) ' Convert to double
        Else
            vlera = 0 ' Default value if invalid
        End If
        If IsNumeric(ws.Cells(i, 5).Value) Then
            detyrim = CDbl(ws.Cells(i, 5).Value) ' Convert to double
        Else
            detyrim = 0 ' Default value if invalid
        End If
        
        ' Duplicate rows based on "piece" count
        For j = 1 To piece
            ws.Cells(newRow, 11).Value = name
            ws.Cells(newRow, 12).Value = piece
            ws.Cells(newRow, 13).Value = vlera
            ws.Cells(newRow, 14).Value = detyrim
            newRow = newRow + 1
        Next j
    Next i
    
    MsgBox "Rows expanded successfully!", vbInformation
End Sub
