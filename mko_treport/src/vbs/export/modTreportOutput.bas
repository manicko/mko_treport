Attribute VB_Name = "modTreportOutput"
Option Explicit

' treport rev.7 - output/report assembly helpers (B9,B10,B11,B3,D1/D2/D7/D11/D12/D14)
' report_sheet placeholders: <max_<PivotTable field name>> and <max-N_<PivotTable field name>>
' resolved per data_files row from the row pivot/data context, independent of data_files.filters
' derivation/representatives. Grammar-aware, exact-token substitution, legacy/malformed/remnant fail-fast.

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

Private Function IsAsciiDigit(ByVal ch As String) As Boolean
    IsAsciiDigit = (Len(ch) = 1 And ch >= "0" And ch <= "9")
End Function

Private Function IsAllAsciiDigits(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String
    If Len(s) = 0 Then
        IsAllAsciiDigits = False
        Exit Function
    End If
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If Not IsAsciiDigit(ch) Then
            IsAllAsciiDigits = False
            Exit Function
        End If
    Next i
    IsAllAsciiDigits = True
End Function

Private Function TryParseMaxPlaceholderToken(ByVal tokenText As String, ByRef outFieldName As String, ByRef outN As Long, ByRef outIsOffset As Boolean) As Boolean
    ' Grammar-aware parse for <max_<field>> and <max<ws>?-<ws>?N<ws>?_<field>>.
    ' max is case-insensitive; N is base-10 [0-9]+, non-negative, no sign/decimal.
    ' Field is exact text between '_' and closing '>' (may include spaces/special chars).
    Dim inner As String
    Dim lowerInner As String
    Dim usPos As Long
    Dim head As String
    Dim fieldPart As String
    Dim i As Long
    Dim rest As String
    Dim nStr As String

    TryParseMaxPlaceholderToken = False
    outFieldName = vbNullString
    outN = 0
    outIsOffset = False

    If Len(tokenText) < 7 Then Exit Function
    If Left$(tokenText, 1) <> "<" Or Right$(tokenText, 1) <> ">" Then Exit Function
    inner = Mid$(tokenText, 2, Len(tokenText) - 2)
    lowerInner = LCase$(inner)
    If Left$(lowerInner, 3) <> "max" Then Exit Function
    usPos = InStr(1, inner, "_", vbBinaryCompare)
    If usPos <= 0 Then Exit Function

    head = Mid$(inner, 1, usPos - 1)
    fieldPart = Mid$(inner, usPos + 1)
    If Len(fieldPart) = 0 Then Exit Function

    If LCase$(head) = "max" Then
        outFieldName = fieldPart
        outN = 0
        outIsOffset = False
        TryParseMaxPlaceholderToken = True
        Exit Function
    End If

    ' Offset form: max<ws>?-<ws>?N<ws>?
    If Len(head) < 5 Then Exit Function ' max-0
    If LCase$(Left$(head, 3)) <> "max" Then Exit Function
    i = 4
    Do While i <= Len(head) And Mid$(head, i, 1) = " "
        i = i + 1
    Loop
    If i > Len(head) Then Exit Function
    If Mid$(head, i, 1) <> "-" Then Exit Function
    i = i + 1
    Do While i <= Len(head) And Mid$(head, i, 1) = " "
        i = i + 1
    Loop
    If i > Len(head) Then Exit Function
    rest = Mid$(head, i)
    ' strip trailing whitespace from N
    Do While Len(rest) > 0 And Right$(rest, 1) = " "
        rest = Left$(rest, Len(rest) - 1)
    Loop
    nStr = rest
    If Not IsAllAsciiDigits(nStr) Then Exit Function

    outFieldName = fieldPart
    outN = CLng(Val(nStr))
    outIsOffset = True
    TryParseMaxPlaceholderToken = True
End Function

Public Function MakeMaxPlaceholderUnitKey(ByVal fieldName As String, ByVal nOffset As Long) As String
    MakeMaxPlaceholderUnitKey = fieldName & vbNullChar & CStr(nOffset)
End Function

Private Function FailMalformedMaxToken(ByVal ctx As String, ByVal template As String, ByVal tokenText As String, ByVal reason As String) As Object
    modTreportConfig.Fail 2614, ctx & ",template='" & template & "',token='" & tokenText & "'", reason
    Set FailMalformedMaxToken = Nothing
End Function

Public Function CollectMaxPlaceholderUnits(ByVal template As String, ByVal ctx As String) As Object
    ' Returns dictionary keyed by (field & vbNullChar & N). Item is dictionary with:
    ' Field (String), N (Long), Tokens (Dictionary of exact token strings -> True), IsOffset (Boolean).
    ' Grammar-aware scanning integrated with legacy/malformed/unresolved/remnant guards.
    Dim units As Object
    Dim s As String
    Dim pos As Long
    Dim ltPos As Long
    Dim gtPos As Long
    Dim tokenText As String
    Dim inner As String
    Dim lowerInner As String
    Dim fieldName As String
    Dim nOffset As Long
    Dim isOffset As Boolean
    Dim unitKey As String
    Dim unit As Object
    Dim tokenMap As Object

    Set units = CreateObject("Scripting.Dictionary")
    units.CompareMode = vbBinaryCompare
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

    pos = 1
    Do
        ltPos = InStr(pos, s, "<", vbBinaryCompare)
        If ltPos = 0 Then Exit Do
        gtPos = InStr(ltPos + 1, s, ">", vbBinaryCompare)
        If gtPos = 0 Then
            ' Unclosed '<...': if it looks like a max placeholder start, fail as malformed max token.
            If InStr(1, LCase$(Mid$(s, ltPos)), "<max", vbBinaryCompare) = 1 Then
                modTreportConfig.Fail 2614, ctx & ",template='" & s & "',token_start=" & CStr(ltPos), "Malformed/unclosed <max_... placeholder token (missing closing '>')."
            End If
            modTreportConfig.Fail 2618, ctx & ",template='" & s & "',token_start=" & CStr(ltPos), "Unresolved placeholder-like token remains in report_sheet template."
        End If
        tokenText = Mid$(s, ltPos, gtPos - ltPos + 1)
        inner = Mid$(tokenText, 2, Len(tokenText) - 2)
        lowerInner = LCase$(inner)

        If Left$(lowerInner, 3) = "max" Then
            If Not TryParseMaxPlaceholderToken(tokenText, fieldName, nOffset, isOffset) Then
                ' Keep deterministic specific cases where possible.
                If lowerInner = "max_" Or Left$(lowerInner, 4) = "max_" Then
                    modTreportConfig.Fail 2615, ctx & ",template='" & s & "',token='" & tokenText & "'", "Malformed <max_... placeholder token: empty or invalid field name."
                End If
                modTreportConfig.Fail 2614, ctx & ",template='" & s & "',token='" & tokenText & "'", "Malformed or unresolved <max...> placeholder token."
            End If

            unitKey = MakeMaxPlaceholderUnitKey(fieldName, nOffset)
            If units.Exists(unitKey) Then
                Set unit = units(unitKey)
                Set tokenMap = unit("Tokens")
                If Not tokenMap.Exists(tokenText) Then tokenMap.Add tokenText, True
            Else
                Set unit = CreateObject("Scripting.Dictionary")
                unit.CompareMode = vbBinaryCompare
                unit.Add "Field", fieldName
                unit.Add "N", nOffset
                unit.Add "IsOffset", isOffset
                Set tokenMap = CreateObject("Scripting.Dictionary")
                tokenMap.CompareMode = vbBinaryCompare
                tokenMap.Add tokenText, True
                unit.Add "Tokens", tokenMap
                units.Add unitKey, unit
            End If
        End If

        pos = gtPos + 1
    Loop

    Set CollectMaxPlaceholderUnits = units
End Function

Public Function CollectMaxPlaceholderFields(ByVal template As String, ByVal ctx As String) As Object
    ' Backward-compatible field set (unique field names) derived from grammar-aware unit collection.
    Dim fields As Object
    Dim units As Object
    Dim k As Variant
    Dim unit As Object
    Dim fieldName As String

    Set fields = CreateObject("Scripting.Dictionary")
    fields.CompareMode = vbBinaryCompare
    Set units = CollectMaxPlaceholderUnits(template, ctx)

    For Each k In units.Keys
        Set unit = units(k)
        fieldName = CStr(unit("Field"))
        If Not fields.Exists(fieldName) Then fields.Add fieldName, True
    Next k

    Set CollectMaxPlaceholderFields = fields
End Function

Public Function SubstituteReportSheetPlaceholders(ByVal template As String, ByVal resolvedMaxValues As Object, ByVal ctx As String) As String
    ' Substitutes all well-formed <max_<field>> / <max-N_<field>> tokens using pre-resolved named-field
    ' values for the processed data_files row (independent of data_files.filters). Exact original token
    ' spellings are substituted; N==0 and base use same resolved value/formatting.
    Dim s As String
    Dim units As Object
    Dim k As Variant
    Dim unit As Object
    Dim fieldName As String
    Dim nOffset As Long
    Dim valKey As String
    Dim valText As String
    Dim tokenMap As Object
    Dim t As Variant
    Dim tokenText As String

    s = template
    Set units = CollectMaxPlaceholderUnits(s, ctx)

    For Each k In units.Keys
        Set unit = units(k)
        fieldName = CStr(unit("Field"))
        nOffset = CLng(unit("N"))
        valKey = MakeMaxPlaceholderUnitKey(fieldName, nOffset)
        Set tokenMap = unit("Tokens")
        For Each t In tokenMap.Keys
            tokenText = CStr(t)
            If resolvedMaxValues Is Nothing Then
                modTreportConfig.Fail 2616, ctx & ",template='" & template & "',field='" & fieldName & "',N=" & CStr(nOffset) & ",token='" & tokenText & "'", "Placeholder '" & tokenText & "' cannot be resolved: no resolved max values provided."
            End If
            If Not resolvedMaxValues.Exists(valKey) Then
                modTreportConfig.Fail 2616, ctx & ",template='" & template & "',field='" & fieldName & "',N=" & CStr(nOffset) & ",token='" & tokenText & "'", "Placeholder '" & tokenText & "' cannot be resolved for referenced PivotTable field '" & fieldName & "' with N=" & CStr(nOffset) & "."
            End If
            valText = CStr(resolvedMaxValues(valKey))
            s = Replace$(s, tokenText, valText)
        Next t
    Next k

    ' RC3: any unresolved/legacy-noncompliant or placeholder-like tokens remaining must fail fast.
    If InStr(1, s, "<max_year>", vbBinaryCompare) > 0 Or InStr(1, s, "<max_monthnum>", vbBinaryCompare) > 0 Or InStr(1, s, "<max_weeknum>", vbBinaryCompare) > 0 Then
        modTreportConfig.Fail 2613, ctx & ",template='" & template & "',result='" & s & "'", "Legacy fixed placeholder token remains unresolved/non-compliant after substitution."
    End If
    If InStr(1, LCase$(s), "<max", vbBinaryCompare) > 0 Then
        modTreportConfig.Fail 2617, ctx & ",template='" & template & "',result='" & s & "'", "Unresolved or malformed <max...> placeholder token remains after substitution."
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




