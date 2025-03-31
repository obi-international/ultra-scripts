' =ROUND((N6 * 99.23) + ((N6 * 99.23) * 0.02) + (((N6 * 99.23) + ((N6 * 99.23) * 0.02)) * 0.2), 2)

Sub ApplyCurrencyConversion()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim sheetName As String, valueColumn As String, resultColumn As String
    Dim exchangeRate As Double
    Dim valueColNum As Integer, resultColNum As Integer
    
    ' Ask user for the sheet name (default: "initially")
    sheetName = InputBox("Enter the sheet name to process:", "Sheet Selection", "initially")
    
    ' Check if the sheet exists
    On Error Resume Next
    Set ws = Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' not found. Please check the name and try again.", vbCritical
        Exit Sub
    End If
    
    ' Ask for the value column (default: "N")
    valueColumn = UCase(InputBox("Enter the column where values are located:", "Value Column Selection", "N"))
    
    ' Convert column letter to number
    valueColNum = Range(valueColumn & "1").Column
    resultColNum = valueColNum + 1  ' The result column is the next column to the right
    resultColumn = Split(Cells(1, resultColNum).Address, "$")(1)  ' Convert column number to letter

    ' Ask for the exchange rate (default: 99.23)
    exchangeRate = CDbl(InputBox("Enter the exchange rate:", "Exchange Rate", 99.23))
    
    ' Find the last row in the value column
    lastRow = ws.Cells(ws.Rows.Count, valueColNum).End(xlUp).Row
    
    ' Loop through each row and apply the formula
    For i = 2 To lastRow  ' Assuming row 1 contains headers
        ws.Cells(i, resultColNum).Formula = "=ROUND(((" & valueColumn & i & " * " & exchangeRate & ") * 0.02) + (((" & valueColumn & i & " * " & exchangeRate & ") + ((" & valueColumn & i & " * " & exchangeRate & ") * 0.02)) * 0.2), 0)"
    Next i
    
    MsgBox "Formula applied successfully in column " & resultColumn, vbInformation
End Sub

