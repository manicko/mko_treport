Attribute VB_Name = "modTreportOutput"
Option Explicit

' treport rev.7 - output/report assembly helpers (B9,B10,B11,D11/D12/D14)
' report_sheet placeholders: <max_<PivotTable field name>> resolved per data_files row
' from the row pivot/data context, independent of data_files.filters derivation/representatives.

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

Public Function CollectMaxPlaceholderFields(ByVal template As String, ByVal ctx As String) As Object
    ' Returns Scripting.Dictionary of unique PivotTable field names referenced by well-formed
    ' <max_<field>> tokens in template (key = field name as token inner text; value = True).
    ' RC3: fails fast on legacy fixed tokens, malformed/unclosed/empty-field <max_... tokens,
    ' and any other unresolved placeholder-like tokens in template.
    Dim fields As Object
    Dim s As String
    Dim pos As Long
    Dim closePos As Long
    Dim innerName As String
    Dim tokenText As String
    Dim prefixPos As Long

    Set fields = CreateObject("Scripting.Dictionary")
    fields.CompareMode = vbBinaryCompare
    s = template

    ' Legacy fixed placeholder tokens are non-compliant (D11/D12, rev.7 migration).
    If InStr(1, s, "<max_year>", vbBinaryCompare) > 0 Then
        modTreportConfig.Fail 2613, ctx & ",template='" & s & "',token='<max_year>'", "Legacy fixed placeholder token '<max_year>' is not allowed; use '<max_<PivotTable field name>>'."
    End If
    If InStr(1, s, "<max_monthnum>", vbBinaryCompare) > 0 Then
        modTreportConfig.Fail 2613, ctx & ",template='" & s & "',token='<max_monthnum>'", "Legacy fixed placeholder token '<max_monthnum>' is not allowed; use '<max_<PivotTable field name>>'."
    End If
    If InStr(1, s, "<max_weeknum>", vbBinaryCompare) > 0 Then
        modTreportConfig.Fail 2613, ctx & ",template='" & s & "',token='<max_weeknum>'", "Legacy fixed placeholder token '<max_weeknum>' is not allowed; use '<max_<PivotTable field name>>'."
    End If

    ' P-Scan: find every '<max_' and the next '>' after it.
    pos = 1
    Do
        prefixPos = InStr(pos, s, "<max_", vbBinaryCompare)
        If prefixPos = 0 Then Exit Do
        closePos = InStr(prefixPos + 5, s, ">", vbBinaryCompare)
        If closePos = 0 Then
            modTreportConfig.Fail 2614, ctx & ",template='" & s & "',token_start=" & CStr(prefixPos), "Malformed/unclosed <max_... placeholder token (missing closing '>')."
        End If
        tokenText = Mid$(s, prefixPos, closePos - prefixPos + 1)
        innerName = Mid$(s, prefixPos + 5, closePos - (prefixPos + 5))
        If Len(innerName) = 0 Then
            modTreportConfig.Fail 2615, ctx & ",template='" & s & "',token='" & tokenText & "'", "Malformed <max_... placeholder token: empty field name."
        End If
        If Not fields.Exists(innerName) Then
            fields.Add innerName, True
        End If
        pos = closePos + 1
    Loop

    ' Any other placeholder-like token in template is unresolved/invalid for report_sheet.
    If InStr(1, s, "<max_", vbBinaryCompare) > 0 And fields.Count = 0 Then
        modTreportConfig.Fail 2614, ctx & ",template='" & s & "'", "Malformed or unresolved <max_... placeholder token."
    End If
    ' Detect placeholder-like tokens not covered by valid <max_<field>> scanning (generic <...> forms).
    ' Valid <max_<field>> tokens are allowed only when they will be resolved from resolvedMaxValues.
    ' Remaining checks are finalized after substitution in SubstituteReportSheetPlaceholders.

    Set CollectMaxPlaceholderFields = fields
End Function

Public Function SubstituteReportSheetPlaceholders(ByVal template As String, ByVal resolvedMaxValues As Object, ByVal ctx As String) As String
    ' Substitutes all well-formed <max_<field>> tokens using pre-resolved named-field maxima for the
    ' processed data_files row (independent of data_files.filters). Fails fast if a referenced field
    ' value is missing/unresolved or if any placeholder-like token remains unresolved/legacy-noncompliant.
    Dim s As String
    Dim fields As Object
    Dim k As Variant
    Dim token As String
    Dim valKey As String
    Dim valText As String

    s = template
    Set fields = CollectMaxPlaceholderFields(s, ctx)

    If Not fields Is Nothing Then
        For Each k In fields.Keys
            token = "<max_" & CStr(k) & ">"
            valKey = CStr(k)
            If resolvedMaxValues Is Nothing Then
                modTreportConfig.Fail 2616, ctx & ",template='" & template & "',field='" & valKey & "',token='" & token & "'", "Placeholder '" & token & "' cannot be resolved: no resolved max values provided."
            End If
            If Not resolvedMaxValues.Exists(valKey) Then
                modTreportConfig.Fail 2616, ctx & ",template='" & template & "',field='" & valKey & "',token='" & token & "'", "Placeholder '" & token & "' cannot be resolved for referenced PivotTable field '" & valKey & "'."
            End If
            valText = CStr(resolvedMaxValues(valKey))
            s = Replace$(s, token, valText)
        Next k
    End If

    ' RC3: any unresolved/legacy-noncompliant or placeholder-like tokens remaining must fail fast.
    If InStr(1, s, "<max_year>", vbBinaryCompare) > 0 Or InStr(1, s, "<max_monthnum>", vbBinaryCompare) > 0 Or InStr(1, s, "<max_weeknum>", vbBinaryCompare) > 0 Then
        modTreportConfig.Fail 2613, ctx & ",template='" & template & "',result='" & s & "'", "Legacy fixed placeholder token remains unresolved/non-compliant after substitution."
    End If
    If InStr(1, s, "<max_", vbBinaryCompare) > 0 Then
        modTreportConfig.Fail 2617, ctx & ",template='" & template & "',result='" & s & "'", "Unresolved or malformed <max_... placeholder token remains after substitution."
    End If
    If HasUnresolvedPlaceholderLikeToken(s) Then
        modTreportConfig.Fail 2618, ctx & ",template='" & template & "',result='" & s & "'", "Unresolved placeholder-like token remains in report_sheet template."
    End If

    SubstituteReportSheetPlaceholders = s
End Function

Private Function HasUnresolvedPlaceholderLikeToken(ByVal s As String) As Boolean
    ' Conservative RC3 guard: any remaining '<...>' token-like fragment is treated as unresolved.
    Dim ltPos As Long
    Dim gtPos As Long
    ltPos = InStr(1, s, "<", vbBinaryCompare)
    If ltPos = 0 Then
        HasUnresolvedPlaceholderLikeToken = False
        Exit Function
    End If
    gtPos = InStr(ltPos + 1, s, ">", vbBinaryCompare)
    HasUnresolvedPlaceholderLikeToken = (gtPos > ltPos)
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

Public Sub CleanupDefaultSheets(ByVal outWb As Workbook, ByVal producedSheetNames As Object, ByVal ctx As String)
    ' Workbook-scoped cleanup for B1/CR-1: remove only unproduced worksheets from wb.Worksheets
    ' on successful normal-completion path before save. Never delete produced report sheets.
    ' producedSheetNames keys are LCase$(Trim$(sheetName)) of final produced sheet names.
    Dim i As Long
    Dim ws As Worksheet
    Dim wsKey As String
    Dim prevAlerts As Boolean
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    If outWb Is Nothing Then
        modTreportConfig.Fail 2620, ctx, "Default-sheet cleanup failed: output workbook is Nothing."
    End If
    If producedSheetNames Is Nothing Then
        modTreportConfig.Fail 2621, ctx, "Default-sheet cleanup failed: produced sheet registry is Nothing."
    End If
    If producedSheetNames.Count = 0 Then
        modTreportConfig.Fail 2622, ctx, "Default-sheet cleanup failed: produced sheet registry is empty (no report/content sheets to keep)."
    End If

    prevAlerts = Application.DisplayAlerts
    On Error GoTo EH
    Application.DisplayAlerts = False

    For i = outWb.Worksheets.Count To 1 Step -1
        Set ws = outWb.Worksheets(i)
        wsKey = LCase$(Trim$(ws.Name))
        If Not producedSheetNames.Exists(wsKey) Then
            If outWb.Worksheets.Count <= 1 Then
                modTreportConfig.Fail 2623, ctx & ",sheet='" & ws.Name & "'", "Default-sheet cleanup cannot remove unproduced sheet because workbook must retain at least one worksheet."
            End If
            ws.Delete
        End If
        Set ws = Nothing
    Next i

    Application.DisplayAlerts = prevAlerts
    On Error GoTo 0
    Exit Sub
EH:
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source
    On Error Resume Next
    Application.DisplayAlerts = prevAlerts
    On Error GoTo 0
    modTreportConfig.Fail 2624, ctx & ",source='" & errSrc & "',error=" & CStr(errNum), "Default-sheet cleanup failed: " & errDesc
End Sub

Public Sub EnsureProducedSheetsPresent(ByVal outWb As Workbook, ByVal producedSheetNames As Object, ByVal ctx As String)
    ' Post-cleanup invariant (B1): intended produced sheets must remain present before save.
    Dim k As Variant
    Dim ws As Worksheet
    Dim found As Boolean
    Dim i As Long
    Dim targetKey As String

    If outWb Is Nothing Then
        modTreportConfig.Fail 2625, ctx, "Output workbook invalid before save: workbook is Nothing."
    End If
    If producedSheetNames Is Nothing Then
        modTreportConfig.Fail 2626, ctx, "Output workbook invalid before save: produced sheet registry is Nothing."
    End If
    If producedSheetNames.Count = 0 Then
        modTreportConfig.Fail 2627, ctx, "Output workbook invalid before save: produced sheet registry is empty."
    End If
    If outWb.Worksheets.Count <= 0 Then
        modTreportConfig.Fail 2628, ctx, "Output workbook invalid before save: no worksheets remain."
    End If

    For Each k In producedSheetNames.Keys
        targetKey = CStr(k)
        found = False
        For i = 1 To outWb.Worksheets.Count
            Set ws = outWb.Worksheets(i)
            If LCase$(Trim$(ws.Name)) = targetKey Then
                found = True
                Exit For
            End If
        Next i
        If Not found Then
            modTreportConfig.Fail 2629, ctx & ",sheet_key='" & targetKey & "'", "Output workbook invalid before save: intended produced sheet not present after cleanup."
        End If
    Next k
End Sub
