Sub FindCombinationsSum()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim colLetter As String
    Dim targetInput As String
    Dim target As Double
    Dim colNum As Long
    
    ' Ask user for sheet name
    sheetName = InputBox("Enter the sheet name:", "Sheet Name", ActiveSheet.Name)
    On Error Resume Next
    Set ws = Sheets(sheetName)
    If ws Is Nothing Then
        MsgBox "Sheet not found.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Ask user for column letter
    colLetter = InputBox("Enter the column letter (e.g., G):", "Column Letter", "G")
    If colLetter = "" Then colLetter = "G"
    colNum = Range(colLetter & "1").Column

    ' Ask user for target value
    targetInput = InputBox("Enter the target sum:", "Target Value", "138.56")
    If Not IsNumeric(targetInput) Then
        MsgBox "Invalid target value.", vbCritical
        Exit Sub
    End If
    target = CDbl(targetInput)

    Dim values() As Double
    Dim lastRow As Long, i As Long
    lastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row
    ReDim values(1 To lastRow)

    For i = 1 To lastRow
        If IsNumeric(ws.Cells(i, colNum).Value) Then
            values(i) = ws.Cells(i, colNum).Value
        Else
            values(i) = 0 ' Treat non-numeric as 0
        End If
    Next i

    ' Loop through combinations
    Dim k As Long, maxCombo As Long
    maxCombo = 2 ' Up to 3-item combinations

    For k = 2 To maxCombo
        Call Combine(values, target, k, 1, "", 0)
    Next k
End Sub

Sub Combine(values() As Double, target As Double, length As Long, start As Long, combo As String, currentSum As Double)
    Dim i As Long
    For i = start To UBound(values)
        Dim newSum As Double
        newSum = currentSum + values(i)
        
        Dim newCombo As String
        newCombo = combo & values(i) & ", "
        
        If Abs(newSum - target) < 0.001 And UBound(Split(newCombo, ",")) >= length Then
            Debug.Print "Match: " & Left(newCombo, Len(newCombo) - 2)
        ElseIf newSum < target And UBound(Split(newCombo, ",")) < length Then
            Call Combine(values, target, length, i + 1, newCombo, newSum)
        End If
    Next i
End Sub
