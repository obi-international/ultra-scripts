Sub CopyGroupedAWBData_NoManifestSum()
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim sourceSheetName As String, targetSheetName As String
    Dim lastRow As Long
    Dim dict As Object
    Dim i As Long
    Dim awb As String, name As String, desc As String
    Dim manifest As Variant
    Dim key As Variant
    Dim rowDest As Long

    ' Prompt for source sheet
    sourceSheetName = InputBox("Enter the sheet name to process:", "Source Sheet Selection", "e")
    Set wsSource = Nothing
    On Error Resume Next
    Set wsSource = Sheets(sourceSheetName)
    On Error GoTo 0
    If wsSource Is Nothing Then
        MsgBox "Sheet '" & sourceSheetName & "' not found.", vbCritical
        Exit Sub
    End If

    ' Prompt for destination sheet
    targetSheetName = InputBox("Enter the sheet name where grouped data will be saved:", "Target Sheet Selection", "edit")
    On Error Resume Next
    Set wsDest = Sheets(targetSheetName)
    On Error GoTo 0
    If wsDest Is Nothing Then
        Set wsDest = Sheets.Add(After:=Sheets(Sheets.Count))
        wsDest.name = targetSheetName
    Else
        wsDest.Cells.Clear
    End If

    ' Set headers
    With wsDest
        .Range("B1").value = "AWB"
        .Range("C1").value = "Marrësi"
        .Range("D1").value = "Përshkrimi"
        .Range("E1").value = "Manifesti"
    End With

    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row

    ' Group data by AWB
    For i = 2 To lastRow
        awb = Trim(wsSource.Cells(i, "B").value)
        name = Trim(wsSource.Cells(i, "C").value)
        desc = Trim(wsSource.Cells(i, "D").value)
        manifest = wsSource.Cells(i, "E").value

        If awb = "" Then GoTo NextIteration

        If Not dict.Exists(awb) Then
            dict.Add awb, Array(name, desc, manifest)
        Else
            Dim arr As Variant
            arr = dict(awb)
            arr(1) = arr(1) & " | " & desc  ' Only concatenate descriptions
            dict(awb) = arr
        End If
NextIteration:
    Next i

    ' Output results
    rowDest = 2
    For Each key In dict.Keys
        Dim valArr As Variant
        valArr = dict(key)
        wsDest.Cells(rowDest, "B").value = key
        wsDest.Cells(rowDest, "C").value = valArr(0)
        wsDest.Cells(rowDest, "D").value = valArr(1)
        wsDest.Cells(rowDest, "E").value = valArr(2) ' Keep first manifest value only
        rowDest = rowDest + 1
    Next key

    MsgBox "AWB data processed in '" & targetSheetName & "' without manifest sum.", vbInformation
End Sub

