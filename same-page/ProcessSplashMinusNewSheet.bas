Sub ProcessAndEditSameSheet()
    Dim sourceWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim processedText As String
    Dim parts As Variant
    Dim sourceSheetName As String

    ' Prompt the user to enter the source sheet name
    sourceSheetName = InputBox("Enter the source sheet name (e.g., original):", "Source Sheet Name", "original")

    ' Set the source worksheet
    On Error Resume Next
    Set sourceWs = ThisWorkbook.Sheets(sourceSheetName)
    On Error GoTo 0
    If sourceWs Is Nothing Then
        MsgBox "Source sheet not found!"
        Exit Sub
    End If

    ' Find the last row in column H
    lastRow = sourceWs.Cells(sourceWs.Rows.Count, 8).End(xlUp).Row ' Column H is the 8th column

    ' Loop through each row in column H of the source sheet
    For i = 2 To lastRow ' Start from row 2 to skip headers
        cellValue = Trim(sourceWs.Cells(i, 8).Value) ' Column H - Trim spaces

        ' Handle slashes `/` or dashes `-`
        If InStr(cellValue, "/") > 0 Or InStr(cellValue, "-") > 0 Then
            parts = Split(cellValue, "/") ' First split by slash
            If UBound(parts) = 0 Then parts = Split(cellValue, "-") ' If no slash, split by dash
            processedText = GetLongestPart(parts)
        Else
            processedText = cellValue
        End If

        ' Convert the processed text to Proper Case
        processedText = Application.WorksheetFunction.Proper(LCase(processedText))

        ' Update the cell in Column H with the processed text
        sourceWs.Cells(i, 8).Value = processedText
    Next i

    MsgBox "Processing complete. Column H has been updated.", vbInformation
End Sub

' Function to get the longest part of the split text
Function GetLongestPart(parts As Variant) As String
    Dim i As Long
    Dim longest As String
    Dim maxLength As Long

    For i = LBound(parts) To UBound(parts)
        If Len(Trim(parts(i))) > maxLength Then
            longest = Trim(parts(i))
            maxLength = Len(Trim(parts(i)))
        End If
    Next i

    GetLongestPart = longest
End Function
