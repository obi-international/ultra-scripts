Sub CopyTotalsByAWBNoNetWeight()
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim sourceSheetName As String, targetSheetName As String
    Dim lastRow As Long
    Dim dict As Object
    Dim i As Long
    Dim awb As String, name As String, city As String
    Dim desc As String, hsCode As String, translation As String
    Dim net As Double, value As Double
    Dim key As Variant
    Dim rowDest As Long

    ' Ask user for the source sheet name
    sourceSheetName = InputBox("Enter the sheet name to process:", "Source Sheet Selection", "init")
    
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
        wsDest.name = targetSheetName
    Else
        wsDest.Cells.Clear
    End If

    ' Set up headers
    With wsDest
        .Range("B1").value = "AWB"
        .Range("C1").value = "Marrësi"
        .Range("D1").value = "Qyteti"
        .Range("E1").value = "Përshkrimi"
        .Range("F1").value = "Net (kg)"
        .Range("G1").value = "Vlera"
        .Range("H1").value = "HS Code"
        .Range("I1").value = "Përkthimi"
    End With

    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row

    ' Collect and group data
    For i = 2 To lastRow
        awb = Trim(wsSource.Cells(i, "B").value)
        name = wsSource.Cells(i, "C").value
        city = wsSource.Cells(i, "D").value
        desc = Trim(wsSource.Cells(i, "E").value)
        hsCode = Trim(wsSource.Cells(i, "H").value)
        translation = Trim(wsSource.Cells(i, "I").value)

        If IsNumeric(wsSource.Cells(i, "F").value) Then
            net = CDbl(wsSource.Cells(i, "F").value)
        Else
            net = 0
        End If
        
        If IsNumeric(wsSource.Cells(i, "G").value) Then
            value = CDbl(wsSource.Cells(i, "G").value)
        Else
            value = 0
        End If

        If awb <> "" Then
            If Not dict.Exists(awb) Then
                dict.Add awb, Array(name, city, desc, net, value, hsCode, translation)
            Else
                Dim arr As Variant
                arr = dict(awb)
                If desc <> "" Then arr(2) = arr(2) & " | " & desc
                If hsCode <> "" Then arr(5) = arr(5) & " | " & hsCode
                If translation <> "" Then arr(6) = arr(6) & " | " & translation
                arr(3) = arr(3) + net
                arr(4) = arr(4) + value
                dict(awb) = arr
            End If
        End If
    Next i

    ' Output to destination sheet
    rowDest = 2
    For Each key In dict.Keys
        Dim valArr As Variant
        valArr = dict(key)

        wsDest.Cells(rowDest, "B").value = key
        wsDest.Cells(rowDest, "C").value = valArr(0)
        wsDest.Cells(rowDest, "D").value = valArr(1)
        wsDest.Cells(rowDest, "E").value = valArr(2)
        wsDest.Cells(rowDest, "F").value = valArr(3) / 1000 ' Convert net to kg
        wsDest.Cells(rowDest, "G").value = valArr(4)
        wsDest.Cells(rowDest, "H").value = valArr(5)
        wsDest.Cells(rowDest, "I").value = valArr(6)
        rowDest = rowDest + 1
    Next key

    MsgBox "Summary complete in sheet '" & targetSheetName & "'!", vbInformation
End Sub
