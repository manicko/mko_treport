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

Private Function FindWorksheetCaseInsensitive(ByVal wb As Workbook, ByVal wsName As String) As Worksheet
    Dim ws As Worksheet
    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, wsName, vbTextCompare) = 0 Then
            Set FindWorksheetCaseInsensitive = ws
            Exit Function
        End If
    Next ws
    Set FindWorksheetCaseInsensitive = Nothing
End Function

Public Function PrepareTargetSheet(ByVal outWb As Workbook, ByVal sheetName As String, ByVal templateName As String, ByVal ctx As String) As Worksheet
    ' P04 GA-7/GA-5/CR-2: prepare target sheet for subreport output.
    ' - templateName blank: create new plain worksheet (ordinary path), preserve existing behavior.
    ' - templateName filled: copy Worksheet template from ThisWorkbook (macro workbook) into outWb.
    ' No existing sheet is overwritten/deleted on conflict (fail-fast). Source template is never modified.
    Dim ws As Worksheet
    Dim existing As Worksheet
    Dim tplWs As Worksheet
    Dim tplNorm As String
    Dim wsCountBefore As Long

    On Error Resume Next
    Set existing = outWb.Worksheets(sheetName)
    On Error GoTo 0
    If Not existing Is Nothing Then
        modTreportConfig.Fail 2634, ctx & ",data_sheet='" & sheetName & "',parameter='template',template='" & templateName & "'", _
            "Target sheet name conflict in output workbook: sheet '" & sheetName & "' already exists and cannot be overwritten or deleted."
    End If

    tplNorm = modTreportConfig.Norm(templateName)
    If Len(tplNorm) = 0 Then
        On Error GoTo EHCreate
        Set ws = outWb.Worksheets.Add(After:=outWb.Worksheets(outWb.Worksheets.Count))
        ws.Name = sheetName
        Set PrepareTargetSheet = ws
        Exit Function
EHCreate:
        modTreportConfig.Fail 2610, ctx & ",data_sheet='" & sheetName & "',parameter='data_sheet/template'", _
            "Unable to create report sheet: " & Err.Description
    End If

    Set tplWs = FindWorksheetCaseInsensitive(ThisWorkbook, tplNorm)
    If tplWs Is Nothing Then
        modTreportConfig.Fail 2635, ctx & ",data_sheet='" & sheetName & "',parameter='template',template='" & tplNorm & "'", _
            "Template worksheet not found or not usable in macro workbook (ThisWorkbook)."
    End If

    On Error GoTo EHCopy
    wsCountBefore = outWb.Worksheets.Count
    tplWs.Copy After:=outWb.Worksheets(outWb.Worksheets.Count)
    If outWb.Worksheets.Count <= wsCountBefore Then
        modTreportConfig.Fail 2636, ctx & ",data_sheet='" & sheetName & "',parameter='template',template='" & tplNorm & "'", _
            "Failed to copy template worksheet into output workbook."
    End If
    Set ws = outWb.Worksheets(outWb.Worksheets.Count)
    On Error GoTo EHCopy

    ws.Name = sheetName
    ws.Visible = xlSheetVisible
    Set PrepareTargetSheet = ws
    Exit Function

EHCopy:
    modTreportConfig.Fail 2636, ctx & ",data_sheet='" & sheetName & "',parameter='template',template='" & tplNorm & "'", _
        "Failed to prepare target sheet from template: " & Err.Description
End Function

Public Function AddReportSheet(ByVal outWb As Workbook, ByVal sheetName As String, ByVal ctx As String) As Worksheet
    ' Compatibility wrapper: ordinary sheet preparation (no template).
    Set AddReportSheet = PrepareTargetSheet(outWb, sheetName, vbNullString, ctx)
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

' =========================
' B2 (P01): naming/group planning and collision preflight
' RC-1..RC-5, D-03/D-04/D-06/OQ-7, CR-01..CR-03/CR-05/CR-06/CR-07/CR-19n..CR-21n/CR-22g/CR-23g
' =========================

Public Type B2SheetMemberPlan
    ReportIndex As Long
    DataFileIndex As Long
    DataFileRowIndex As Long
    EmissionOrder As Long
    GapRows As Long
    EffectiveBaseName As String
    GroupKey As String
    PlannedName As String
    SuffixText As String
    SuffixLength As Long
    OutputName As String
    ReportName As String
    FileCtx As String
    RepCtx As String
    RawReportSheetTemplate As String
End Type

Public Type B2SheetGroupPlan
    GroupKey As String
    EffectiveBaseName As String
    MemberCount As Long
    Members() As B2SheetMemberPlan
End Type

Public Function NormalizeSheetNameChars(ByVal nameIn As String) As String
    ' D-03/CR-01: apply existing sanitization character rules without length validation.
    Dim s As String, i As Long, ch As String, bad As String
    s = modTreportConfig.Norm(nameIn)
    bad = ":\/?*[]"
    For i = 1 To Len(bad)
        ch = Mid$(bad, i, 1)
        s = Replace$(s, ch, "_")
    Next i
    s = Replace$(s, "'", "_")
    NormalizeSheetNameChars = s
End Function

Public Function ResolveEffectiveBaseSheetName(ByVal substitutedName As String, ByVal ctx As String, ByVal rawTemplateOptional As String) As String
    ' Resolve placeholders/substitutions upstream, then normalize/sanitize to effective base name (no numeric suffix).
    ' 2601 preserved for empty; legacy 2602 is not used for planning-validated D-04 cases.
    Dim baseName As String
    baseName = NormalizeSheetNameChars(substitutedName)
    If Len(baseName) = 0 Then
        modTreportConfig.Fail 2601, ctx & B2RawTemplateContext(rawTemplateOptional), "Resulting sheet name is empty after normalization."
    End If
    ResolveEffectiveBaseSheetName = baseName
End Function

Private Function B2RawTemplateContext(ByVal rawTemplateOptional As String) As String
    If Len(rawTemplateOptional) > 0 Then
        B2RawTemplateContext = ",raw_report_sheet_template='" & rawTemplateOptional & "'"
    Else
        B2RawTemplateContext = vbNullString
    End If
End Function

Public Function MakeSheetGroupKey(ByVal effectiveBaseName As String) As String
    ' D-03/CR-01: case-insensitive grouping key within the same output workbook.
    MakeSheetGroupKey = LCase$(effectiveBaseName)
End Function

Private Function B2PlanningContext(ByRef m As B2SheetMemberPlan) As String
    Dim c As String
    c = m.RepCtx
    If Len(m.OutputName) > 0 Then c = c & ",output='" & m.OutputName & "'"
    If m.ReportIndex > 0 Then c = c & ",report_index=" & CStr(m.ReportIndex)
    If Len(m.ReportName) > 0 Then c = c & ",report_name='" & m.ReportName & "'"
    If m.DataFileIndex > 0 Then c = c & ",data_file_index=" & CStr(m.DataFileIndex)
    If m.DataFileRowIndex > 0 Then c = c & ",data_file_row=" & CStr(m.DataFileRowIndex)
    If Len(m.RawReportSheetTemplate) > 0 Then c = c & ",raw_report_sheet_template='" & m.RawReportSheetTemplate & "'"
    B2PlanningContext = c
End Function

Private Sub FailPlannedSheetNameLength(ByRef m As B2SheetMemberPlan)
    ' 2637 (D-04/CR-19n..21n): planning-validated worksheet-name length violation.
    ' Required: effective base name (without numeric suffix), explicit max permitted base length (31 - suffix_length),
    ' suffix_length/suffix and concrete planned name, output/report/data_file/row context; raw template optional.
    Dim maxBaseLen As Long
    Dim msg As String
    maxBaseLen = 31 - m.SuffixLength
    If maxBaseLen < 0 Then maxBaseLen = 0
    msg = "Planned worksheet name exceeds 31 characters (D-04). " & _
          "Effective base name after substitution/sanitization without numeric suffix: '" & m.EffectiveBaseName & "'. " & _
          "Maximum permitted base-name length (31 - suffix_length): " & CStr(maxBaseLen) & ". " & _
          "suffix_length=" & CStr(m.SuffixLength) & ", suffix='" & m.SuffixText & "'. " & _
          "Concrete planned worksheet name: '" & m.PlannedName & "'."
    modTreportConfig.Fail 2637, B2PlanningContext(m), msg
End Sub

Public Sub AssignPlannedSheetNames(ByRef members() As B2SheetMemberPlan)
    ' OQ-7/CR-05/D-04: base for first member, then base1..base(N-1) in original member order.
    ' Uses actual suffix length when validating (suffix_length >= 0, including multi-digit).
    Dim i As Long
    Dim lb As Long, ub As Long
    lb = LBound(members)
    ub = UBound(members)
    For i = lb To ub
        If i = lb Then
            members(i).SuffixText = vbNullString
            members(i).SuffixLength = 0
            members(i).PlannedName = members(i).EffectiveBaseName
        Else
            members(i).SuffixText = CStr(i - lb)
            members(i).SuffixLength = Len(members(i).SuffixText)
            members(i).PlannedName = members(i).EffectiveBaseName & members(i).SuffixText
        End If
        If Len(members(i).PlannedName) > 31 Then
            FailPlannedSheetNameLength members(i)
        End If
    Next i
End Sub

Public Sub BuildOutputSheetPlan(ByRef members() As B2SheetMemberPlan, ByRef groupsByKey As Object, ByRef plannedNames As Object)
    ' Builds per-output grouping plan and all planned-name map in original member order.
    ' groupsByKey: key=LCase effective base -> Dictionary with Keys: Key, BaseName, Members(Collection in order)
    ' plannedNames: key=LCase planned name -> Dictionary info for diagnostics/preflight.
    Dim i As Long
    Dim gk As String
    Dim grp As Object
    Dim coll As Object
    Dim info As Object
    Dim pnKey As String

    For i = LBound(members) To UBound(members)
        gk = members(i).GroupKey
        If Not groupsByKey.Exists(gk) Then
            Set grp = CreateObject("Scripting.Dictionary")
            grp.CompareMode = vbBinaryCompare
            grp.Add "Key", gk
            grp.Add "BaseName", members(i).EffectiveBaseName
            Set coll = CreateObject("Scripting.Dictionary")
            coll.CompareMode = vbBinaryCompare
            grp.Add "Members", coll
            groupsByKey.Add gk, grp
        Else
            Set grp = groupsByKey(gk)
            Set coll = grp("Members")
        End If
        Set coll = grp("Members")
        coll.Add CStr(coll.Count + 1), CloneB2MemberInfo(members(i))

        pnKey = LCase$(members(i).PlannedName)
        Set info = CreateObject("Scripting.Dictionary")
        info.CompareMode = vbBinaryCompare
        info.Add "PlannedName", members(i).PlannedName
        info.Add "GroupKey", gk
        info.Add "BaseName", members(i).EffectiveBaseName
        info.Add "Context", B2PlanningContext(members(i))
        info.Add "MemberOrder", CStr(members(i).EmissionOrder)
        If plannedNames.Exists(pnKey) Then
            ' Duplicate planned name in current run: fail as collision preflight (2638).
            FailPlannedNameCollision members(i).PlannedName, "planned in current run", CStr(plannedNames(pnKey)("Context")), _
                                     "planned in current run", B2PlanningContext(members(i))
        End If
        plannedNames.Add pnKey, info
    Next i
End Sub

Private Function CloneB2MemberInfo(ByRef m As B2SheetMemberPlan) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbBinaryCompare
    d.Add "ReportIndex", m.ReportIndex
    d.Add "DataFileIndex", m.DataFileIndex
    d.Add "DataFileRowIndex", m.DataFileRowIndex
    d.Add "EmissionOrder", m.EmissionOrder
    d.Add "GapRows", m.GapRows
    d.Add "EffectiveBaseName", m.EffectiveBaseName
    d.Add "GroupKey", m.GroupKey
    d.Add "PlannedName", m.PlannedName
    d.Add "SuffixText", m.SuffixText
    d.Add "SuffixLength", m.SuffixLength
    d.Add "OutputName", m.OutputName
    d.Add "ReportName", m.ReportName
    d.Add "FileCtx", m.FileCtx
    d.Add "RepCtx", m.RepCtx
    d.Add "RawReportSheetTemplate", m.RawReportSheetTemplate
    Set CloneB2MemberInfo = d
End Function

Private Sub FailPlannedNameCollision(ByVal conflictingName As String, ByVal classA As String, ByVal ctxA As String, ByVal classB As String, ByVal ctxB As String)
    ' 2638 (D-06/CR-06/22g): concrete collision preflight.
    Dim msg As String
    msg = "Worksheet name collision detected before creating affected worksheets (D-06). " & _
          "Conflicting worksheet name: '" & conflictingName & "'. " & _
          "Conflict class: " & classA & " vs " & classB & ". " & _
          "Context A: " & ctxA & ". Context B: " & ctxB & "."
    modTreportConfig.Fail 2638, ctxA & "; " & ctxB, msg
End Sub

Public Sub PreflightPlannedSheetNameCollisions(ByVal outWb As Workbook, ByVal plannedNames As Object, ByVal outCtx As String)
    ' D-06/CR-06/CR-22g: check existing worksheet names and all planned run names before creating affected worksheets.
    ' Case-insensitive. Never overwrite/delete/rename.
    Dim ws As Worksheet
    Dim existingKey As String
    Dim planKey As Variant
    Dim plannedName As String
    Dim planCtx As String

    If outWb Is Nothing Then
        modTreportConfig.Fail 2638, outCtx, "Worksheet name collision preflight failed: output workbook is Nothing."
    End If

    ' Existing workbook worksheets vs planned names.
    For Each ws In outWb.Worksheets
        existingKey = LCase$(ws.Name)
        If plannedNames.Exists(existingKey) Then
            plannedName = CStr(plannedNames(existingKey)("PlannedName"))
            planCtx = CStr(plannedNames(existingKey)("Context"))
            modTreportConfig.Fail 2638, outCtx & ",existing_sheet='" & ws.Name & "'; " & planCtx, _
                "Worksheet name collision detected before creating affected worksheets (D-06). " & _
                "Conflicting worksheet name: '" & ws.Name & "'. " & _
                "Conflict class: existing workbook worksheet vs planned in current run. " & _
                "Planned context: " & planCtx & "."
        End If
    Next ws

    ' Planned-vs-planned duplicates are detected during BuildOutputSheetPlan (2638).
    ' Keep an explicit pass for defensive completeness (no-op if unique).
    For Each planKey In plannedNames.Keys
        plannedName = CStr(plannedNames(planKey)("PlannedName"))
        planCtx = CStr(plannedNames(planKey)("Context"))
        If Len(plannedName) = 0 Then
            modTreportConfig.Fail 2638, outCtx & "; " & planCtx, "Worksheet name collision preflight failed: planned worksheet name is empty."
        End If
    Next planKey
End Sub

Public Function HasCombineGroups(ByVal groupsByKey As Object) As Boolean
    Dim k As Variant
    Dim grp As Object
    Dim coll As Object
    For Each k In groupsByKey.Keys
        Set grp = groupsByKey(k)
        Set coll = grp("Members")
        If coll.Count > 1 Then
            HasCombineGroups = True
            Exit Function
        End If
    Next k
    HasCombineGroups = False
End Function

Public Function DescribeCombineGroups(ByVal groupsByKey As Object, ByVal outputName As String) As String
    ' RC-1 diagnostic helper: summarize multi-member groups for fail-fast context.
    Dim k As Variant
    Dim grp As Object
    Dim coll As Object
    Dim m As Variant
    Dim mi As Object
    Dim parts As String
    Dim detail As String
    Dim first As Boolean
    first = True
    parts = vbNullString
    For Each k In groupsByKey.Keys
        Set grp = groupsByKey(k)
        Set coll = grp("Members")
        If coll.Count > 1 Then
            detail = "[output='" & outputName & "',effective_base='" & CStr(grp("BaseName")) & "',group_key='" & CStr(grp("Key")) & "',member_count=" & CStr(coll.Count) & ",members="
            Dim j As Long
            For j = 1 To coll.Count
                Set mi = coll(CStr(j))
                detail = detail & "{order=" & CStr(mi("EmissionOrder")) & ",df_row=" & CStr(mi("DataFileRowIndex")) & ",report='" & CStr(mi("ReportName")) & "',planned='" & CStr(mi("PlannedName")) & "'}"
                If j < coll.Count Then detail = detail & ";"
            Next j
            detail = detail & "]"
            If first Then
                parts = detail
                first = False
            Else
                parts = parts & " | " & detail
            End If
        End If
    Next k
    DescribeCombineGroups = parts
End Function

Public Function GetPlannedSheetNameForMember(ByRef members() As B2SheetMemberPlan, ByVal emissionOrder As Long) As String
    Dim i As Long
    For i = LBound(members) To UBound(members)
        If members(i).EmissionOrder = emissionOrder Then
            GetPlannedSheetNameForMember = members(i).PlannedName
            Exit Function
        End If
    Next i
    GetPlannedSheetNameForMember = vbNullString
End Function

' =========================
' B5 (P01): combined-group finalization — temp deletion and final/base registry reconciliation
' CR-18c, D-06/CR-22g. Runs only after successful B4 combine per group, before cleanup/verify/save.
' Fail-fast: 2698 for precondition/invariant/registry/state-resolution (no destructive side effects for detectable incoherence),
' 2696 for delete/verify/COM failures. No temp deletion before successful B4; no temps in finalProducedNames; no save/cleanup/verify.
' =========================

Private Function B5DictHas(ByVal d As Object, ByVal key As String) As Boolean
    On Error Resume Next
    If d Is Nothing Then
        B5DictHas = False
    Else
        B5DictHas = d.Exists(key)
    End If
    On Error GoTo 0
End Function

Private Function B5DictStr(ByVal d As Object, ByVal key As String, ByVal defaultValue As String) As String
    On Error Resume Next
    If B5DictHas(d, key) Then
        B5DictStr = CStr(d(key))
    Else
        B5DictStr = defaultValue
    End If
    If Err.Number <> 0 Then B5DictStr = defaultValue
    On Error GoTo 0
End Function

Private Function B5DictLong(ByVal d As Object, ByVal key As String, ByVal defaultValue As Long) As Long
    On Error Resume Next
    If B5DictHas(d, key) Then
        B5DictLong = CLng(d(key))
    Else
        B5DictLong = defaultValue
    End If
    If Err.Number <> 0 Then B5DictLong = defaultValue
    On Error GoTo 0
End Function

Private Function B5DictBool(ByVal d As Object, ByVal key As String, ByVal defaultValue As Boolean) As Boolean
    On Error Resume Next
    If B5DictHas(d, key) Then
        B5DictBool = CBool(d(key))
    Else
        B5DictBool = defaultValue
    End If
    If Err.Number <> 0 Then B5DictBool = defaultValue
    On Error GoTo 0
End Function

Private Function B5NormName(ByVal s As String) As String
    B5NormName = LCase$(Trim$(s))
End Function

Private Function B5SafeWsName(ByVal ws As Worksheet) As String
    On Error Resume Next
    If ws Is Nothing Then
        B5SafeWsName = "<Nothing>"
    Else
        B5SafeWsName = ws.Name
    End If
    If Err.Number <> 0 Then B5SafeWsName = "<unavailable>"
    On Error GoTo 0
End Function

Private Function B5WorksheetBelongsToWorkbook(ByVal wb As Workbook, ByVal ws As Worksheet) As Boolean
    Dim parentWb As Workbook
    B5WorksheetBelongsToWorkbook = False
    If wb Is Nothing Or ws Is Nothing Then Exit Function
    On Error Resume Next
    Set parentWb = ws.Parent
    If Err.Number = 0 Then
        If Not parentWb Is Nothing Then
            B5WorksheetBelongsToWorkbook = (parentWb Is wb)
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function B5ResolveWorksheetInWorkbook(ByVal wb As Workbook, ByVal mi As Object, ByVal plannedName As String) As Worksheet
    Dim ws As Worksheet
    Dim wsByName As Worksheet

    Set B5ResolveWorksheetInWorkbook = Nothing
    If wb Is Nothing Then Exit Function

    ' Prefer runtime worksheet reference only when it is a valid sheet in the expected workbook.
    Set ws = Nothing
    On Error Resume Next
    If B5DictHas(mi, "Sheet") Then Set ws = mi("Sheet")
    If Err.Number <> 0 Then Set ws = Nothing
    Err.Clear
    On Error GoTo 0
    If Not ws Is Nothing Then
        If B5WorksheetBelongsToWorkbook(wb, ws) Then
            If Len(plannedName) = 0 Or StrComp(ws.Name, plannedName, vbTextCompare) = 0 Then
                Set B5ResolveWorksheetInWorkbook = ws
                Exit Function
            End If
            ' Runtime ref is in workbook but name differs; allow fallback by planned name if consistent object exists.
        Else
            Set ws = Nothing
        End If
    End If

    If Len(plannedName) > 0 Then
        Set wsByName = FindWorksheetCaseInsensitive(wb, plannedName)
        If Not wsByName Is Nothing Then
            Set B5ResolveWorksheetInWorkbook = wsByName
            Exit Function
        End If
    End If

    ' Last resort: runtime ref in workbook even if name differs and planned lookup failed.
    If Not ws Is Nothing Then
        If B5WorksheetBelongsToWorkbook(wb, ws) Then Set B5ResolveWorksheetInWorkbook = ws
    End If
End Function

Private Function B5CommonGroupCtx(ByVal ctx As String, ByVal groupRecord As Object, ByVal outputName As String) As String
    B5CommonGroupCtx = ctx & ",output='" & outputName & "'" & _
        ",effective_base='" & B5DictStr(groupRecord, "BaseName", vbNullString) & "'" & _
        ",group_key='" & B5DictStr(groupRecord, "GroupKey", vbNullString) & "'"
End Function

Private Function B5MemberCtx(ByVal mi As Object, ByVal groupRecord As Object, ByVal memberIndex As Long, ByVal resolvedName As String) As String
    Dim s As String
    s = ",member_index=" & CStr(memberIndex) & _
        ",is_base=" & CStr(B5DictBool(mi, "IsBase", (memberIndex = 1))) & _
        ",is_temp=" & CStr(B5DictBool(mi, "IsTemp", (memberIndex > 1))) & _
        ",prepared=" & CStr(B5DictBool(mi, "Prepared", False)) & _
        ",planned_sheet='" & B5DictStr(mi, "PlannedName", vbNullString) & "'" & _
        ",member_group_key='" & B5DictStr(mi, "GroupKey", vbNullString) & "'" & _
        ",emission_order=" & CStr(B5DictLong(mi, "EmissionOrder", 0)) & _
        ",report_index=" & CStr(B5DictLong(mi, "ReportIndex", 0)) & _
        ",data_file_index=" & CStr(B5DictLong(mi, "DataFileIndex", 0)) & _
        ",data_file_row=" & CStr(B5DictLong(mi, "DataFileRowIndex", 0))
    If Len(resolvedName) > 0 Then s = s & ",actual_sheet_name='" & resolvedName & "'"
    B5MemberCtx = s
End Function

Private Sub B5FailInvariant(ByVal ctx As String, ByVal msg As String)
    modTreportConfig.Fail 2698, ctx & ",operation='b5_finalize'", msg
End Sub

Private Sub B5FailDelete(ByVal ctx As String, ByVal msg As String)
    modTreportConfig.Fail 2696, ctx & ",operation='b5_finalize'", msg
End Sub

Public Sub FinalizeCombinedGroupTemps(ByVal outBook As Workbook, ByVal groupRecord As Object, ByVal finalProducedNames As Object, ByVal ctx As String)
    ' B5: delete temp worksheets only after successful B4 combine and reconcile final/base registry/runtime state.
    ' Postcondition on success: base/final remains present and in finalProducedNames; temps absent from workbook and never registered.
    ' On failure: fail-fast (2698 invariant/state, 2696 delete/verify/COM) before cleanup/verify/save; no false-success state implied.
    Dim outputName As String
    Dim groupCtx As String
    Dim membersColl As Object
    Dim memberCount As Long
    Dim j As Long
    Dim baseMi As Object
    Dim mi As Object
    Dim baseWs As Worksheet
    Dim basePlannedName As String
    Dim baseKey As String
    Dim tempKeys() As String
    Dim tempNames() As String
    Dim tempSheets() As Worksheet
    Dim tempMembers() As Object
    Dim tempCount As Long
    Dim tempMi As Object
    Dim tempWs As Worksheet
    Dim tempPlanned As String
    Dim tempKey As String
    Dim tempActualName As String
    Dim prevAlerts As Variant
    Dim alertsSuppressed As Boolean
    Dim delCtx As String
    Dim errNum As Long
    Dim errSrc As String
    Dim errDesc As String

    alertsSuppressed = False
    On Error GoTo EH

    ' ----------------------------
    ' Phase A: validate inputs/state and resolve all targets (NO destructive actions in this phase)
    ' ----------------------------
    If outBook Is Nothing Then B5FailInvariant ctx & ",invariant='outbook_nothing'", "B5 finalization failed: output workbook is Nothing."
    If groupRecord Is Nothing Then B5FailInvariant ctx & ",invariant='group_nothing'", "B5 finalization failed: group record is Nothing."
    If finalProducedNames Is Nothing Then B5FailInvariant ctx & ",invariant='final_registry_nothing'", "B5 finalization failed: final produced registry is Nothing."

    If Not B5DictHas(groupRecord, "Members") Then
        B5FailInvariant ctx & ",invariant='members_missing'", "B5 finalization failed: group Members collection is missing."
    End If
    On Error Resume Next
    Set membersColl = groupRecord("Members")
    On Error GoTo EH
    If membersColl Is Nothing Then
        B5FailInvariant ctx & ",invariant='members_nothing'", "B5 finalization failed: group Members collection is Nothing."
    End If
    On Error Resume Next
    memberCount = membersColl.Count
    If Err.Number <> 0 Then memberCount = 0
    Err.Clear
    On Error GoTo EH
    If memberCount < 2 Then
        B5FailInvariant ctx & ",invariant='member_count',member_count=" & CStr(memberCount) & ",effective_base='" & B5DictStr(groupRecord, "BaseName", vbNullString) & "'", _
            "B5 finalization failed: combine group requires at least 2 prepared members."
    End If

    Set baseMi = Nothing
    If membersColl.Exists("1") Then Set baseMi = membersColl("1")
    If baseMi Is Nothing Then
        B5FailInvariant ctx & ",invariant='base_member_missing',member_index=1", "B5 finalization failed: base member (index 1) is missing."
    End If

    outputName = B5DictStr(baseMi, "OutputNameRuntime", B5DictStr(baseMi, "OutputName", vbNullString))
    If Len(outputName) = 0 Then outputName = B5DictStr(groupRecord, "OutputName", vbNullString)
    groupCtx = B5CommonGroupCtx(ctx, groupRecord, outputName)

    If Not B5DictBool(baseMi, "IsBase", False) Then
        B5FailInvariant groupCtx & B5MemberCtx(baseMi, groupRecord, 1, vbNullString) & ",invariant='base_flag'", "B5 finalization failed: member 1 is not marked as base."
    End If
    If B5DictBool(baseMi, "IsTemp", False) Then
        B5FailInvariant groupCtx & B5MemberCtx(baseMi, groupRecord, 1, vbNullString) & ",invariant='base_temp_flag'", "B5 finalization failed: member 1 is marked as temp."
    End If
    If Not B5DictBool(baseMi, "Prepared", False) Then
        B5FailInvariant groupCtx & B5MemberCtx(baseMi, groupRecord, 1, vbNullString) & ",invariant='base_not_prepared'", "B5 finalization failed: base member is not marked Prepared."
    End If

    basePlannedName = B5DictStr(baseMi, "PlannedName", vbNullString)
    If Len(basePlannedName) = 0 Then
        B5FailInvariant groupCtx & B5MemberCtx(baseMi, groupRecord, 1, vbNullString) & ",invariant='base_planned_name_empty'", "B5 finalization failed: base PlannedName is empty."
    End If
    Set baseWs = B5ResolveWorksheetInWorkbook(outBook, baseMi, basePlannedName)
    If baseWs Is Nothing Then
        B5FailInvariant groupCtx & B5MemberCtx(baseMi, groupRecord, 1, vbNullString) & ",invariant='base_sheet_missing',planned_sheet='" & basePlannedName & "'", _
            "B5 finalization failed: base worksheet not found/resolvable in output workbook after successful combine."
    End If
    If Not B5WorksheetBelongsToWorkbook(outBook, baseWs) Then
        B5FailInvariant groupCtx & B5MemberCtx(baseMi, groupRecord, 1, B5SafeWsName(baseWs)) & ",invariant='base_sheet_wrong_workbook'", _
            "B5 finalization failed: resolved base worksheet does not belong to target output workbook."
    End If

    baseKey = B5NormName(baseWs.Name)
    If Len(baseKey) = 0 Then
        B5FailInvariant groupCtx & B5MemberCtx(baseMi, groupRecord, 1, B5SafeWsName(baseWs)) & ",invariant='base_key_empty'", "B5 finalization failed: base worksheet name key is empty."
    End If

    ' Validate and collect temp members (2..N) before any destructive action.
    tempCount = memberCount - 1
    ReDim tempKeys(1 To tempCount)
    ReDim tempNames(1 To tempCount)
    ReDim tempSheets(1 To tempCount)
    ReDim tempMembers(1 To tempCount)

    For j = 2 To memberCount
        Set tempMi = Nothing
        If membersColl.Exists(CStr(j)) Then Set tempMi = membersColl(CStr(j))
        If tempMi Is Nothing Then
            B5FailInvariant groupCtx & ",member_index=" & CStr(j) & ",invariant='temp_member_missing'", "B5 finalization failed: temp member record is missing."
        End If
        Set tempMembers(j - 1) = tempMi

        If B5DictBool(tempMi, "IsBase", False) Then
            B5FailInvariant groupCtx & B5MemberCtx(tempMi, groupRecord, j, vbNullString) & ",invariant='temp_marked_base'", "B5 finalization failed: non-first member is marked as base."
        End If
        If Not B5DictBool(tempMi, "IsTemp", True) Then
            B5FailInvariant groupCtx & B5MemberCtx(tempMi, groupRecord, j, vbNullString) & ",invariant='temp_flag_missing'", "B5 finalization failed: non-first member is not marked as temp."
        End If
        If Not B5DictBool(tempMi, "Prepared", False) Then
            B5FailInvariant groupCtx & B5MemberCtx(tempMi, groupRecord, j, vbNullString) & ",invariant='temp_not_prepared'", "B5 finalization failed: temp member is not marked Prepared."
        End If

        tempPlanned = B5DictStr(tempMi, "PlannedName", vbNullString)
        If Len(tempPlanned) = 0 Then
            B5FailInvariant groupCtx & B5MemberCtx(tempMi, groupRecord, j, vbNullString) & ",invariant='temp_planned_name_empty'", "B5 finalization failed: temp PlannedName is empty."
        End If
        tempKey = B5NormName(tempPlanned)
        If Len(tempKey) = 0 Then
            B5FailInvariant groupCtx & B5MemberCtx(tempMi, groupRecord, j, vbNullString) & ",invariant='temp_key_empty'", "B5 finalization failed: temp worksheet name key is empty."
        End If
        If StrComp(tempKey, baseKey, vbBinaryCompare) = 0 Or StrComp(tempKey, B5NormName(basePlannedName), vbBinaryCompare) = 0 Then
            B5FailInvariant groupCtx & B5MemberCtx(tempMi, groupRecord, j, vbNullString) & ",invariant='temp_same_as_base',base_key='" & baseKey & "',temp_key='" & tempKey & "'", _
                "B5 finalization failed: temp worksheet identity collides with base/final worksheet."
        End If
        If finalProducedNames.Exists(tempKey) Then
            B5FailInvariant groupCtx & B5MemberCtx(tempMi, groupRecord, j, vbNullString) & ",invariant='temp_registered_as_final',registry_key='" & tempKey & "'", _
                "B5 finalization failed: temp worksheet is registered in final produced registry; temps must not be final-produced."
        End If

        Set tempWs = B5ResolveWorksheetInWorkbook(outBook, tempMi, tempPlanned)
        If tempWs Is Nothing Then
            B5FailInvariant groupCtx & B5MemberCtx(tempMi, groupRecord, j, vbNullString) & ",invariant='temp_sheet_missing',planned_sheet='" & tempPlanned & "',base_key='" & baseKey & "'", _
                "B5 finalization failed: temp worksheet not found/resolvable in output workbook after successful combine; strict invariant requires temp presence before deletion."
        End If
        If Not B5WorksheetBelongsToWorkbook(outBook, tempWs) Then
            B5FailInvariant groupCtx & B5MemberCtx(tempMi, groupRecord, j, B5SafeWsName(tempWs)) & ",invariant='temp_sheet_wrong_workbook',planned_sheet='" & tempPlanned & "'", _
                "B5 finalization failed: resolved temp worksheet does not belong to target output workbook."
        End If
        If tempWs Is baseWs Then
            B5FailInvariant groupCtx & B5MemberCtx(tempMi, groupRecord, j, B5SafeWsName(tempWs)) & ",invariant='temp_resolves_to_base',planned_sheet='" & tempPlanned & "'", _
                "B5 finalization failed: resolved temp worksheet target is the base/final worksheet."
        End If

        tempActualName = B5SafeWsName(tempWs)
        tempNames(j - 1) = tempActualName
        tempKeys(j - 1) = B5NormName(tempActualName)
        Set tempSheets(j - 1) = tempWs

        If StrComp(B5NormName(tempActualName), baseKey, vbBinaryCompare) = 0 Then
            B5FailInvariant groupCtx & B5MemberCtx(tempMi, groupRecord, j, tempActualName) & ",invariant='temp_name_equals_base',base_key='" & baseKey & "'", _
                "B5 finalization failed: temp worksheet name resolves to base/final worksheet."
        End If

        Set tempWs = Nothing
        Set tempMi = Nothing
    Next j

    ' Base/final registry reconciliation (idempotent, base-only). No temp registration/removal of finals.
    If Not finalProducedNames.Exists(baseKey) Then
        finalProducedNames.Add baseKey, True
    End If

    ' ----------------------------
    ' Phase B: delete temp worksheets (post-B4 successful transfer, post-Phase-A validation)
    ' ----------------------------
    On Error GoTo EH
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    alertsSuppressed = True

    For j = 1 To tempCount
        Set tempWs = tempSheets(j)
        tempActualName = tempNames(j)
        Set tempMi = tempMembers(j)

        delCtx = groupCtx & B5MemberCtx(tempMi, groupRecord, j + 1, tempActualName) & ",base_actual_sheet_name='" & B5SafeWsName(baseWs) & "',base_key='" & baseKey & "',temp_registry_key='" & tempKeys(j) & "'"

        If tempWs Is Nothing Then
            GoTo DeleteVerifyOrStateFail
        End If
        If Not B5WorksheetBelongsToWorkbook(outBook, tempWs) Then
            Set tempWs = FindWorksheetCaseInsensitive(outBook, tempActualName)
            Set tempSheets(j) = tempWs
        End If
        If tempWs Is Nothing Then
DeleteVerifyOrStateFail:
            errNum = 0: errSrc = vbNullString: errDesc = "temp worksheet target is Nothing before delete"
            GoTo DeleteFail
        End If
        If tempWs Is baseWs Then
            errNum = 0: errSrc = vbNullString: errDesc = "resolved temp worksheet target is base/final worksheet"
            GoTo DeleteFail
        End If
        If StrComp(B5NormName(B5SafeWsName(tempWs)), baseKey, vbBinaryCompare) = 0 Then
            errNum = 0: errSrc = vbNullString: errDesc = "temp worksheet name matches base/final worksheet before delete"
            GoTo DeleteFail
        End If
        If outBook.Worksheets.Count <= 1 Then
            errNum = 0: errSrc = vbNullString: errDesc = "cannot delete worksheet because output workbook would have no remaining sheets"
            GoTo DeleteFail
        End If

        errNum = 0: errSrc = vbNullString: errDesc = vbNullString
        On Error Resume Next
        Err.Clear
        tempWs.Delete
        errNum = Err.Number
        errSrc = Err.Source
        errDesc = Err.Description
        On Error GoTo EH
        If errNum <> 0 Then GoTo DeleteFail

        ' Verify removal immediately (case-insensitive).
        Set tempWs = FindWorksheetCaseInsensitive(outBook, tempActualName)
        If Not tempWs Is Nothing Then
            errNum = 0: errSrc = "B5Finalize"
            errDesc = "temp worksheet still present after delete verification"
            Set tempSheets(j) = tempWs
            GoTo DeleteFail
        End If

        Set tempSheets(j) = Nothing
        Set tempWs = Nothing
    Next j

    Application.DisplayAlerts = prevAlerts
    alertsSuppressed = False

    ' ----------------------------
    ' Phase C: runtime reconciliation on full success only
    ' ----------------------------
    On Error GoTo EH
    For j = 1 To tempCount
        Set tempMi = tempMembers(j)
        If Not tempMi Is Nothing Then
            If tempMi.Exists("Sheet") Then Set tempMi("Sheet") = Nothing Else tempMi.Add "Sheet", Nothing
            If tempMi.Exists("InsertedRange") Then Set tempMi("InsertedRange") = Nothing
            If tempMi.Exists("CombineStatus") Then tempMi("CombineStatus") = "TempDeleted" Else tempMi.Add "CombineStatus", "TempDeleted"
        End If
        Set tempMi = Nothing
    Next j

    ' Keep base runtime reference valid if available; run sets group CombineStatus=Finalized after successful hook return.
    If Not baseMi Is Nothing Then
        If baseMi.Exists("Sheet") Then Set baseMi("Sheet") = baseWs Else baseMi.Add "Sheet", baseWs
        If baseMi.Exists("CombineStatus") Then baseMi("CombineStatus") = "BaseFinal" Else baseMi.Add "CombineStatus", "BaseFinal"
    End If

    Exit Sub

DeleteFail:
    If alertsSuppressed Then
        On Error Resume Next
        Application.DisplayAlerts = prevAlerts
        alertsSuppressed = False
        On Error GoTo 0
    End If
    modTreportConfig.Fail 2696, delCtx & _
        ",err_number=" & CStr(errNum) & _
        ",err_source='" & Replace$(errSrc, "'", "''") & _
        "',err_description='" & Replace$(errDesc, "'", "''") & "'", _
        "B5 combined-group finalization failed: temp worksheet deletion/verification did not complete. " & _
        "Group/base context: " & delCtx & ". " & _
        "Temp deletion is allowed only after successful B4 transfer; finalization aborted before save (output workbook must not be saved). " & _
        "COM error: " & CStr(errNum) & " / " & errSrc & " / " & errDesc

EH:
    If alertsSuppressed Then
        On Error Resume Next
        Application.DisplayAlerts = prevAlerts
        alertsSuppressed = False
        On Error GoTo 0
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
