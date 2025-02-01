Sub GenerateArrayFromColumnM()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim outputText As String
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Find the last used row in column M
    lastRow = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row
    
    ' Start the array definition
    outputText = "names = Array( _" & vbCrLf & "    "
    
    ' Loop through each name in column M
    For i = 1 To lastRow
        outputText = outputText & """" & ws.Cells(i, "M").Value & """, "
        
        ' Add a line break after every 10 names for readability
        If i Mod 10 = 0 Then
            outputText = outputText & " _" & vbCrLf & "    "
        End If
    Next i
    
    ' Remove the trailing comma and space
    If Right(outputText, 2) = ", " Then
        outputText = Left(outputText, Len(outputText) - 2)
    End If
    
    ' Close the array definition
    outputText = outputText & " _" & vbCrLf & ")"
    
    ' Write the result to cell A1
    ws.Range("A1").Value = outputText
End Sub
