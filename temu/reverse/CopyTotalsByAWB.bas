Sub CopyTotalsByAWB()
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim sourceSheetName As String, targetSheetName As String
    Dim lastRow As Long
    Dim dict As Object
    Dim i As Long
    Dim awb As String, name As String, city As String, desc As String
    Dim net As Double, value As Double
    Dim key As Variant
    Dim rowDest As Long

    ' Ask user for the source sheet name
    sourceSheetName = InputBox("Enter the sheet name to process:", "Source Sheet Selection", "e")
    
    On Error Resume Next
    Set wsSource = Sheets(sourceSheetName)
    On Error GoTo 0
    If wsSource Is Nothing Then
        MsgBox "Sheet '" & sourceSheetName & "' not found. Please check the name and try again.", vbCritical
        Exit Sub
    End If

    ' Ask user for the target sheet name
    targetSheetName = InputBox("Enter the sheet name where filtered rows will be saved:", "Target Sheet Selection", "edit")
    
    On Error Resume Next
    Set wsDest = Sheets(targetSheetName)
    On Error GoTo 0
    If wsDest Is Nothing Then
        Set wsDest = Sheets.Add(After:=Sheets(Sheets.Count))
        wsDest.Name = targetSheetName
    Else
        wsDest.Cells.Clear
    End If

    ' Set up headers
    With wsDest
        .Range("B1").Value = "AWB"
        .Range("C1").Value = "Marrësi"
        .Range("D1").Value = "Qyteti"
        .Range("E1").Value = "Përshkrimi"
        .Range("F1").Value = "Net"
        .Range("G1").Value = "Vlera"
    End With

    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row

    ' Collect and group data
    For i = 2 To lastRow
        awb = wsSource.Cells(i, "B").Value
        name = wsSource.Cells(i, "C").Value
        city = wsSource.Cells(i, "D").Value
        desc = wsSource.Cells(i, "E").Value
        net = wsSource.Cells(i, "F").Value
        value = wsSource.Cells(i, "G").Value

        If Not dict.Exists(awb) Then
            dict.Add awb, Array(name, city, desc, CDbl(net), CDbl(value))
        Else
            Dim arr As Variant
            arr = dict(awb)
            arr(2) = arr(2) & " | " & desc     ' concatenate descriptions
            arr(3) = arr(3) + CDbl(net)        ' sum net
            arr(4) = arr(4) + CDbl(value)      ' sum value
            dict(awb) = arr
        End If
    Next i

    ' Output to destination sheet
    rowDest = 2
    For Each key In dict.Keys
        Dim valArr As Variant
        valArr = dict(key)
        wsDest.Cells(rowDest, "B").Value = key           ' AWB
        wsDest.Cells(rowDest, "C").Value = valArr(0)     ' Name
        wsDest.Cells(rowDest, "D").Value = valArr(1)     ' City
        wsDest.Cells(rowDest, "E").Value = valArr(2)     ' Description
        wsDest.Cells(rowDest, "F").Value = valArr(3)     ' Net
        wsDest.Cells(rowDest, "G").Value = valArr(4)     ' Value
        rowDest = rowDest + 1
    Next key

    MsgBox "Summary complete in sheet '" & targetSheetName & "'!", vbInformation
End Sub
