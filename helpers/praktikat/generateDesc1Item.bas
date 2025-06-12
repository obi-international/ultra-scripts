Sub GenerateDescriptionsOneItem()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim lastRow As Long, i As Long
    Dim fullName As String, numPieces As String, importNum As String
    Dim detyrimi As Variant, totPieces As String
    Dim description As String

    ' Prompt for the sheet name
    sheetName = InputBox("Enter the sheet name:", "Sheet Name", "teke")

    ' Check if the sheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Sheet name is incorrect. Please check and try again.", vbCritical
        Exit Sub
    End If

    ' Find the last row in column D
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row

    ' Loop through each row and generate the description
    For i = 2 To lastRow ' Assuming headers are in row 1
        fullName = UCase(ws.Cells(i, "B").Value)
        numPieces = ws.Cells(i, "C").Value
        totPieces = ws.Cells(i, "J").Value
        detyrimi = ws.Cells(i, "G").Value
        importNum = ws.Cells(i, "H").Value

        ' Construct the description text
        description = fullName & vbNewLine & _
                      numPieces & " PAKO DERGESE POSTARE " & importNum & vbNewLine & _
                      totPieces & " " & detyrimi

        ' Write the description in column I of the same sheet
        ws.Cells(i, "I").Value = description
    Next i

    MsgBox "Descriptions generated successfully in column I of " & sheetName, vbInformation
End Sub
