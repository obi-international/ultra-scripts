Sub GenerateDescriptions()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim lastRow As Long, i As Long
    Dim fullName As String, numPieces As String, importNum As String
    Dim detyrimi As Variant ' Use Variant to handle both text and numbers
    Dim description As String
    
    ' Prompt for the sheet name
    sheetName = InputBox("Enter the sheet name:", "Sheet Name", "ready")
    
    ' Check if the sheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet name is incorrect. Please check and try again.", vbCritical
        Exit Sub
    End If
    
    ' Find the last row in column F
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    
    ' Loop through each row and generate the description
    For i = 2 To lastRow ' Assuming headers are in row 1
        fullName = UCase(ws.Cells(i, "B").Value)
        numPieces = ws.Cells(i, "D").Value
        importNum = ws.Cells(i, "I").Value
        detyrimi = ws.Cells(i, "H").Value
        
        ' Convert detyrimi to a number if possible
        If IsNumeric(detyrimi) Then
            detyrimi = Application.WorksheetFunction.Round(CDbl(detyrimi), 0)
        Else
            detyrimi = "0" ' Fallback in case of non-numeric value
        End If
        
        ' Construct the description text
        description = fullName & vbNewLine & _
                      numPieces & " PAKO DERGESA POSTARE " & importNum & vbNewLine & _
                      "D-" & detyrimi
        
        ' Write the description in column U of the same sheet
        ws.Cells(i, "J").Value = description
    Next i
    
    MsgBox "Descriptions generated successfully in column J of " & sheetName, vbInformation
End Sub




