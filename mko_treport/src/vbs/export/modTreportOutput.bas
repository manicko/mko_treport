Attribute VB_Name = "modTreportOutput"
Option Explicit

' treport rev.5 - output/report assembly helpers (B9,B10,B11)

Public Function SanitizeSheetName(ByVal nameIn As String, ByVal ctx As String) As String
    Dim s As String, i As Long, ch As String, bad As String
    s = modTreportConfig.Norm(nameIn)
    bad = ":\/?*[]"
    For i = 1 To Len(bad)
        ch = Mid$(bad, i, 1)
        s = Replace$(s, ch, "_")
    Next i
    s = Replace$(s, "'", "_")
    If Len(s) = 0 Then modTreportConfig.Fail 2601, ctx, "Resulting sheet name is empty after normalization."
    If Len(s) > 31 Then modTreportConfig.Fail 2602, ctx, "Resulting sheet name exceeds 31 characters: '" & s & "'"
    SanitizeSheetName = s
End Function

Public Function SanitizeFileName(ByVal nameIn As String, ByVal ctx As String) As String
    Dim s As String, i As Long, ch As String, bad As String
    s = modTreportConfig.Norm(nameIn)
    bad = "\/:*?""<>|"
    For i = 1 To Len(bad)
        ch = Mid$(bad, i, 1)
        s = Replace$(s, ch, "_")
    Next i
    If Len(s) = 0 Then modTreportConfig.Fail 2603, ctx, "Resulting report file name is empty after normalization."
    SanitizeFileName = s
End Function

Public Sub EnsureNoSheetNameCollision(ByVal wb As Workbook, ByVal sheetName As String, ByVal ctx As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If Not ws Is Nothing Then
        modTreportConfig.Fail 2604, ctx, "Sheet name collision after normalization: '" & sheetName & "'."
    End If
End Sub

Public Function SubstitutePlaceholders(ByVal template As String, ByVal yearVal As Variant, ByVal monthVal As Variant, ByVal weekVal As Variant, ByVal ctx As String) As String
    Dim s As String
    s = template ' keep internal spaces; Trim applied by caller on final name as needed
    If InStr(1, s, "<max_year>", vbBinaryCompare) > 0 Then
        If IsEmpty(yearVal) Or IsNull(yearVal) Then modTreportConfig.Fail 2605, ctx, "Placeholder <max_year> referenced but value cannot be derived."
        s = Replace$(s, "<max_year>", CStr(CLng(CDbl(yearVal))))
    End If
    If InStr(1, s, "<max_monthnum>", vbBinaryCompare) > 0 Then
        If IsEmpty(monthVal) Or IsNull(monthVal) Then modTreportConfig.Fail 2606, ctx, "Placeholder <max_monthnum> referenced but value cannot be derived."
        s = Replace$(s, "<max_monthnum>", CStr(CLng(CDbl(monthVal))))
    End If
    If InStr(1, s, "<max_weeknum>", vbBinaryCompare) > 0 Then
        If IsEmpty(weekVal) Or IsNull(weekVal) Then modTreportConfig.Fail 2607, ctx, "Placeholder <max_weeknum> referenced but value cannot be derived."
        s = Replace$(s, "<max_weeknum>", CStr(CLng(CDbl(weekVal))))
    End If
    SubstitutePlaceholders = s
End Function

Public Function GetOrCreateOutputWorkbook(ByRef outWb As Workbook, ByVal outPath As String, ByVal fileNameNoExt As String, ByVal ctx As String) As Workbook
    If Not outWb Is Nothing Then
        Set GetOrCreateOutputWorkbook = outWb
        Exit Function
    End If
    On Error GoTo EH
    Set outWb = Application.Workbooks.Add(xlWBATWorksheet)
    Set GetOrCreateOutputWorkbook = outWb
    Exit Function
EH:
    modTreportConfig.Fail 2608, ctx & ",file='" & outPath & fileNameNoExt & ".xlsx'", "Unable to create output workbook: " & Err.Description
End Function

Public Sub SaveOutputWorkbook(ByRef outWb As Workbook, ByVal outFolder As String, ByVal fileNameNoExt As String, ByVal ctx As String)
    Dim fullPath As String
    If outWb Is Nothing Then Exit Sub
    fullPath = outFolder & fileNameNoExt & ".xlsx"
    On Error GoTo EH
    Application.DisplayAlerts = False
    outWb.SaveAs fileName:=fullPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    outWb.Close SaveChanges:=False
    Set outWb = Nothing
    Exit Sub
EH:
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not outWb Is Nothing Then outWb.Close SaveChanges:=False
    On Error GoTo 0
    Set outWb = Nothing
    modTreportConfig.Fail 2609, ctx & ",file='" & fullPath & "'", "Unable to save output workbook: " & Err.Description
End Sub

Public Function AddReportSheet(ByVal outWb As Workbook, ByVal sheetName As String, ByVal ctx As String) As Worksheet
    Dim ws As Worksheet, existing As Worksheet
    On Error Resume Next
    Set existing = outWb.Worksheets(sheetName)
    On Error GoTo 0
    If Not existing Is Nothing Then
        Application.DisplayAlerts = False
        On Error Resume Next
        existing.Delete
        On Error GoTo 0
        Application.DisplayAlerts = True
    End If
    On Error GoTo EH
    Set ws = outWb.Worksheets.Add(After:=outWb.Worksheets(outWb.Worksheets.Count))
    ws.Name = sheetName
    ' remove default sheet later if extra; keep simple
    Set AddReportSheet = ws
    Exit Function
EH:
    modTreportConfig.Fail 2610, ctx & ",sheet='" & sheetName & "'", "Unable to create report sheet: " & Err.Description
End Function

Public Sub CleanupDefaultSheets(ByVal outWb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    For Each ws In outWb.Worksheets
        If outWb.Worksheets.Count <= 1 Then Exit For
        If Left$(ws.Name, 5) = "Sheet" Or ws.Name = "Лист1" Then
            ' delete only if empty-ish and not a named report target; safer to leave alone to avoid deleting real sheets
        End If
    Next ws
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub
