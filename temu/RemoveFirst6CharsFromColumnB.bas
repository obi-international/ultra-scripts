Sub TrimCharsInColumnC()
    Dim wsName As String
    Dim ws As Worksheet
    Dim charsToRemove As Variant
    Dim r As Long
    Dim cellVal As String

    ' Prompt for sheet name (default "date")
    wsName = InputBox("Enter the sheet name:", "Sheet Selection", "date")
    If wsName = "" Then Exit Sub

    ' Prompt for number of characters to remove (default 3)
    charsToRemove = InputBox("Enter number of characters to remove:", "Trim Settings", 6)
    If Not IsNumeric(charsToRemove) Or charsToRemove < 0 Then
        MsgBox "Please enter a valid number of characters to remove.", vbExclamation
        Exit Sub
    End If

    ' Get the worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(wsName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet '" & wsName & "' not found.", vbCritical
        Exit Sub
    End If

    ' Start from row 20 in column C
    r = 20
    Do While ws.Cells(r, "C").Value <> ""
        cellVal = ws.Cells(r, "C").Value
        If Len(cellVal) > charsToRemove Then
            ws.Cells(r, "C").Value = Mid(cellVal, charsToRemove + 1)
        Else
            ws.Cells(r, "C").Value = "" ' Clear if text is shorter than chars to remove
        End If
        r = r + 1
    Loop

    MsgBox "Trimming complete in column C starting from row 20.", vbInformation
End Sub

