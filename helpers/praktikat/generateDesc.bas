Sub GenerateDescriptions()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim lastRow As Long, i As Long
    Dim fullName As String, numPieces As String, importNum As String
    Dim referenceNum As Variant ' Use Variant to handle both text and numbers
    Dim description As String
    
    ' Prompt for the sheet name
    sheetName = InputBox("Enter the sheet name:","Sheet Name", "ready")
    
    ' Check if the sheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet name is incorrect. Please check and try again.", vbCritical
        Exit Sub
    End If
    
    ' Find the last row in column F
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    
    ' Loop through each row and generate the description
    For i = 2 To lastRow ' Assuming headers are in row 1
        fullName = ws.Cells(i, "F").Value
        numPieces = ws.Cells(i, "N").Value
        importNum = ws.Cells(i, "T").Value
        referenceNum = ws.Cells(i, "R").Value
        
        ' Convert referenceNum to a number if possible
        If IsNumeric(referenceNum) Then
            referenceNum = Application.WorksheetFunction.Round(CDbl(referenceNum), 0)
        Else
            referenceNum = "0" ' Fallback in case of non-numeric value
        End If
        
        ' Construct the description text
        description = fullName & vbNewLine & _
                      numPieces & " pako dergese postare " & importNum & vbNewLine & _
                      "D-" & referenceNum
        
        ' Write the description in column U of the same sheet
        ws.Cells(i, "U").Value = description
    Next i
    
    MsgBox "Descriptions generated successfully in column U of " & sheetName, vbInformation
End Sub


