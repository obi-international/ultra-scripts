Sub GroupAndCopyTotals()
    Dim ws As Worksheet, wsNew As Worksheet
    Dim lastRow As Long, i As Long, outputRow As Long
    Dim nameDict As Object
    Dim sourceSheet As String, destSheet As String

    ' Prompt for source and destination sheet names
    sourceSheet = InputBox("Enter the name of the source sheet:", "Source Sheet", "tot")
    If sourceSheet = "" Then Exit Sub
    destSheet = InputBox("Enter the name of the final sheet:", "Destination Sheet", "ready")
    If destSheet = "" Then Exit Sub

    ' Check if source sheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sourceSheet)
    If ws Is Nothing Then
        MsgBox "The source sheet '" & sourceSheet & "' does not exist!", vbCritical, "Error"
        Exit Sub
    End If
    On Error GoTo 0

    Set nameDict = CreateObject("Scripting.Dictionary")

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Column A = name

    ' Loop through each row
    For i = 2 To lastRow ' skip header
        Dim name As String: name = ws.Cells(i, 1).Value
        Dim place As String: place = ws.Cells(i, 2).Value
        Dim piece As Double: piece = ws.Cells(i, 3).Value
        Dim neto As Double: neto = ws.Cells(i, 4).Value
        Dim bruto As Double: bruto = ws.Cells(i, 5).Value
        Dim value As Double: value = ws.Cells(i, 6).Value

        If Not nameDict.exists(name) Then
            ' place, piece, neto, bruto, value
            nameDict.Add name, Array(place, 0, 0, 0, 0)
        End If

        Dim totals As Variant
        totals = nameDict(name)
        totals(1) = totals(1) + piece
        totals(2) = totals(2) + neto
        totals(3) = totals(3) + bruto
        totals(4) = totals(4) + value
        nameDict(name) = totals
    Next i

    ' Create or clear destination sheet
    On Error Resume Next
    Set wsNew = ThisWorkbook.Sheets(destSheet)
    If wsNew Is Nothing Then
        Set wsNew = ThisWorkbook.Sheets.Add
        wsNew.Name = destSheet
    Else
        wsNew.Cells.Clear
    End If
    On Error GoTo 0

    ' Write headers
    wsNew.Cells(1, 1).Value = "Name"
    wsNew.Cells(1, 2).Value = "Place"
    wsNew.Cells(1, 3).Value = "Total Piece"
    wsNew.Cells(1, 4).Value = "Total Neto"
    wsNew.Cells(1, 5).Value = "Total Bruto"
    wsNew.Cells(1, 6).Value = "Total Value"
    wsNew.Rows(1).Font.Bold = True

    ' Output grouped results
    outputRow = 2
    Dim key As Variant
    For Each key In nameDict.keys
        wsNew.Cells(outputRow, 1).Value = key
        wsNew.Cells(outputRow, 2).Value = nameDict(key)(0)
        wsNew.Cells(outputRow, 3).Value = nameDict(key)(1)
        wsNew.Cells(outputRow, 4).Value = nameDict(key)(2)
        wsNew.Cells(outputRow, 5).Value = nameDict(key)(3)
        wsNew.Cells(outputRow, 6).Value = nameDict(key)(4)
        outputRow = outputRow + 1
    Next key

    MsgBox "Data successfully grouped and copied to '" & destSheet & "' sheet!", vbInformation, "Done"
End Sub
