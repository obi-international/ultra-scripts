' ColA	ColB	ColC:Marresi	COP	Pesha Bruto	PESHE Neto	VLERA	PERSHKRIMI	CODE	Sasia ColJ:Artikujve

Sub GroupByHSCodeWithHeadersAndTotals_Conditional()
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim sourceSheetName As String
    Dim destSheetName As String
    Dim lastRow As Long
    Dim i As Long, destRow As Long
    Dim codeDict As Object
    Dim cell As Range
    Dim code As String
    Dim r As Range
    Dim key As Variant
    Dim totalCO As Double, totalBruto As Double, totalNeto As Double, totalSasi As Double

    ' Ask user for the source sheet name
    sourceSheetName = InputBox("Enter the sheet name to copy from:", "Source Sheet", "Sheet1")
    If sourceSheetName = "" Then Exit Sub

    ' Delete existing "HS CODES" sheet if it exists
    destSheetName = "HS CODES"
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(destSheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Set worksheets
    Set wsSource = ThisWorkbook.Sheets(sourceSheetName)
    Set wsDest = ThisWorkbook.Sheets.Add
    wsDest.Name = destSheetName

    Set codeDict = CreateObject("Scripting.Dictionary")

    ' Get the last row of source data
    lastRow = wsSource.Cells(wsSource.Rows.Count, "I").End(xlUp).Row

    ' Build a dictionary of CODEs and associated rows
    For i = 2 To lastRow
        code = Trim(wsSource.Cells(i, "I").Value)
        If Not codeDict.exists(code) Then
            Set codeDict(code) = New Collection
        End If
        codeDict(code).Add wsSource.Rows(i)
    Next i

    ' Start writing from row 1
    destRow = 1

    ' Loop through each unique CODE
    For Each key In codeDict.Keys
        ' Add a heading for the group
        wsDest.Cells(destRow, 1).Value = "CODE: " & key
        wsDest.Cells(destRow, 1).Font.Bold = True
        destRow = destRow + 1

        ' Copy header row
        wsSource.Rows(1).Copy Destination:=wsDest.Rows(destRow)
        destRow = destRow + 1

        ' Reset totals
        totalCO = 0
        totalBruto = 0
        totalNeto = 0
        totalSasi = 0
        totalValue = 0

        ' Copy rows and accumulate totals
        For Each r In codeDict(key)
            r.Copy Destination:=wsDest.Rows(destRow)
            totalCO = totalCO + Val(r.Cells(1, "D").Value)
            totalBruto = totalBruto + Val(r.Cells(1, "E").Value)
            totalNeto = totalNeto + Val(r.Cells(1, "F").Value)
            totalValue = totalValue + Val(r.Cells(1, "G").Value)
            totalSasi = totalSasi + Val(r.Cells(1, "J").Value)
            destRow = destRow + 1
        Next r

        ' Insert TOTAL row only if more than one data row
        If codeDict(key).Count > 1 Then
            With wsDest
                .Cells(destRow, "C").Value = "TOTAL"
                .Cells(destRow, "D").Value = totalCO
                .Cells(destRow, "E").Value = totalBruto
                .Cells(destRow, "F").Value = totalNeto
                .Cells(destRow, "G").Value = totalValue
                .Cells(destRow, "J").Value = totalSasi
                .Range(.Cells(destRow, "C"), .Cells(destRow, "J")).Font.Bold = True
            End With
            destRow = destRow + 1
        End If

        ' Add two blank rows after each group
        destRow = destRow + 1
    Next key

    MsgBox "Data grouped and totals added (if applicable) in '" & destSheetName & "'.", vbInformation
End Sub
