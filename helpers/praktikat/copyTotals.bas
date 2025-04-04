Sub GroupAndCopyTotals()
    Dim ws As Worksheet, wsNew As Worksheet
    Dim lastRow As Long, i As Long, outputRow As Long
    Dim nameDict As Object
    Dim sourceSheet As String, destSheet As String
    
    ' Prompt for source sheet name
    sourceSheet = InputBox("Enter the name of the source sheet:", "Source Sheet", "initially")
    If sourceSheet = "" Then Exit Sub ' Exit if no input
    
    ' Prompt for destination sheet name
    destSheet = InputBox("Enter the name of the final sheet:", "Destination Sheet", "ready")
    If destSheet = "" Then Exit Sub ' Exit if no input

    ' Check if source sheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sourceSheet)
    If ws Is Nothing Then
        MsgBox "The source sheet '" & sourceSheet & "' does not exist!", vbCritical, "Error"
        Exit Sub
    End If
    On Error GoTo 0
    
    Set nameDict = CreateObject("Scripting.Dictionary")
    
    ' Find last row with data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Assuming "Name" is in column A
    
    ' Loop through each row and accumulate totals
    For i = 2 To lastRow ' Assuming headers are in row 1
        Dim name As String: name = ws.Cells(i, 1).Value
        Dim location As String: location = ws.Cells(i, 2).Value
        Dim phone As String: phone = ws.Cells(i, 3).Value
        Dim pieces As Double: pieces = ws.Cells(i, 4).Value
        Dim weight1 As Double: weight1 = ws.Cells(i, 5).Value
        Dim weight2 As Double: weight2 = ws.Cells(i, 6).Value
        Dim value1 As Double: value1 = ws.Cells(i, 7).Value
        
        ' Store values in dictionary
        If Not nameDict.exists(name) Then
            nameDict.Add name, Array(location, phone, 0, 0, 0, 0) ' Initialize (Location, Phone, Pieces, Weight1, Weight2, Value1)
        End If
        
        Dim totals As Variant
        totals = nameDict(name)
        totals(2) = totals(2) + pieces
        totals(3) = totals(3) + weight1
        totals(4) = totals(4) + weight2
        totals(5) = totals(5) + value1
        nameDict(name) = totals
    Next i

    ' Create or clear the destination sheet
    On Error Resume Next
    Set wsNew = ThisWorkbook.Sheets(destSheet)
    If wsNew Is Nothing Then
        Set wsNew = ThisWorkbook.Sheets.Add
        wsNew.name = destSheet
    Else
        wsNew.Cells.Clear ' Clear existing data
    End If
    On Error GoTo 0

    ' Write header to destination sheet
    wsNew.Cells(1, 1).Value = "Name"
    wsNew.Cells(1, 2).Value = "Location"
    wsNew.Cells(1, 3).Value = "Phone"
    wsNew.Cells(1, 4).Value = "Total Pieces"
    wsNew.Cells(1, 5).Value = "Total Weight1"
    wsNew.Cells(1, 6).Value = "Total Weight2"
    wsNew.Cells(1, 7).Value = "Total Value1"
    wsNew.Rows(1).Font.Bold = True
    
    ' Output grouped totals
    outputRow = 2
    Dim key As Variant
    For Each key In nameDict.keys
        wsNew.Cells(outputRow, 1).Value = key
        wsNew.Cells(outputRow, 2).Value = nameDict(key)(0)
        wsNew.Cells(outputRow, 3).Value = nameDict(key)(1)
        wsNew.Cells(outputRow, 4).Value = nameDict(key)(2)
        wsNew.Cells(outputRow, 5).Value = nameDict(key)(3)
        wsNew.Cells(outputRow, 6).Value = nameDict(key)(4)
        wsNew.Cells(outputRow, 7).Value = nameDict(key)(5)
        outputRow = outputRow + 1
    Next key

    MsgBox "Data successfully grouped and copied to '" & destSheet & "' sheet!", vbInformation, "Done"
End Sub

