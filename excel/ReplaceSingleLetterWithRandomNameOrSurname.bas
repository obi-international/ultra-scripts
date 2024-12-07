Sub ReplaceSingleLetterWithRandomNameOrSurname()
    Dim sourceWs As Worksheet, targetWs As Worksheet
    Dim lastRow As Long, targetRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim splitParts As Variant
    Dim surnames As Variant, names As Variant
    Dim randomName As String, randomSurname As String
    Dim sourceSheetName As String, targetSheetName As String
    
    ' Prompt the user to enter the source and target sheet names
    sourceSheetName = InputBox("Enter the source sheet name (e.g., perpunuar):", "Source Sheet Name", "perpunuar")
    targetSheetName = InputBox("Enter the target sheet name (e.g., perpunuar.):", "Target Sheet Name", "perpunuar.")
    
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
    
    ' Create or clear the target worksheet
    On Error Resume Next
    Set targetWs = ThisWorkbook.Sheets(targetSheetName)
    If targetWs Is Nothing Then
        Set targetWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        targetWs.Name = targetSheetName
    Else
        targetWs.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Copy headers if applicable (e.g., first row)
    sourceWs.Rows(1).Copy Destination:=targetWs.Rows(1)
    
    ' Initialize target row
    targetRow = 2 ' Assuming the first row is headers
    
    ' Define the list of random surnames and names
    surnames = Array("Hoxha", "Paja", "Dajti", "Balliu", "Shala", "Hoxhaj", "Leka", "Kaci", "Muca")
    names = Array("Arti", "Blerina", "Ermal", "Flora", "Gentian", "Ilir", "Jonida", "Kristi", "Luan", "Mira")
    
    ' Loop through each row in column H of the source sheet
    For i = 2 To lastRow ' Start from row 2 to skip headers
        cellValue = Trim(sourceWs.Cells(i, 8).Value) ' Column H - Trim spaces
        
        ' Split the cell value by space
        splitParts = Split(cellValue, " ")
        
        ' Loop through parts of the cell value
        Dim processedParts As String
        processedParts = ""
        Dim part As Variant
        For Each part In splitParts
            If Len(part) = 1 Then
                ' Replace single letters with a random name or surname
                If processedParts = "" Then
                    ' If it's the first part, replace with a name
                    randomName = names(Int((UBound(names) + 1) * Rnd))
                    processedParts = processedParts & " " & randomName
                Else
                    ' Otherwise, replace with a surname
                    randomSurname = surnames(Int((UBound(surnames) + 1) * Rnd))
                    processedParts = processedParts & " " & randomSurname
                End If
            Else
                ' Keep the part as is if it's not a single letter
                processedParts = processedParts & " " & part
            End If
        Next part
        
        ' Remove leading/trailing spaces from the processed text
        processedParts = Trim(processedParts)
        
        ' Write the processed text to the target sheet
        sourceWs.Rows(i).Copy Destination:=targetWs.Rows(targetRow)
        targetWs.Cells(targetRow, 8).Value = processedParts
        
        ' Increment target row counter
        targetRow = targetRow + 1
    Next i
End Sub
