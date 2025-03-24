Sub NumberRedBackgroundRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim colorCode As Long
    Dim counter As Integer
    Dim startRow As Long
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Define red background color used for highlighting (RGB 255, 0, 0)
    colorCode = RGB(255, 0, 0)
    
    ' Find the last row in column F
    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row ' Column F is the 6th column
    
    ' Start from row 19
    startRow = 3
    counter = 1 ' Initialize counter
    
    ' Loop through column F from row 19 and below
    For Each cell In ws.Range("F" & startRow & ":F" & lastRow)
        ' Check if the row is highlighted in red
        If cell.Interior.Color = colorCode Then
            ' Place the counter value in column E of the same row
            ws.Cells(cell.Row, 5).Value = counter
            ' Increment counter
            counter = counter + 1
        End If
    Next cell
    
    MsgBox "Numbering of red-highlighted rows completed in column E.", vbInformation
End Sub
