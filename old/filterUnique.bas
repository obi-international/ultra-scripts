Sub FilterUniqueByMobile()
    Dim sourceSheetName As String
    Dim targetSheetName As String
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim lastRow As Long
    Dim dict As Object
    Dim i As Long
    Dim nextRow As Long

    ' Prompt for sheet names
    sourceSheetName = InputBox("Enter the source sheet name:", "Source Sheet", "e")
    targetSheetName = InputBox("Enter the target sheet name:", "Target Sheet", "filtered")
    
    ' Set source sheet
    On Error Resume Next
    Set sourceSheet = ThisWorkbook.Sheets(sourceSheetName)
    On Error GoTo 0
    If sourceSheet Is Nothing Then
        MsgBox "Source sheet '" & sourceSheetName & "' not found.", vbCritical
        Exit Sub
    End If
    
    ' Delete existing target sheet if exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(targetSheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Create new target sheet
    Set targetSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    targetSheet.Name = targetSheetName
    
    ' Initialize dictionary
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Get last row in source sheet
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "D").End(xlUp).Row
    
    ' Copy headers
    sourceSheet.Rows(1).Copy Destination:=targetSheet.Rows(1)
    nextRow = 2
    
    ' Loop through rows to check for unique mobiles
    For i = 2 To lastRow
        Dim mobile As String
        mobile = Trim(sourceSheet.Cells(i, 4).Value)
        
        If Len(mobile) > 0 Then
            If Not dict.exists(mobile) Then
                dict.Add mobile, True
                sourceSheet.Rows(i).Copy Destination:=targetSheet.Rows(nextRow)
                nextRow = nextRow + 1
            End If
        End If
    Next i
    
    MsgBox "Unique rows copied to '" & targetSheetName & "'.", vbInformation
End Sub

