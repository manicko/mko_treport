Attribute VB_Name = "modTreportFilters"
Option Explicit

' Generic filter parsing/classification/validation (T1/B2) and generic per-field min/max resolution/apply (T2/T3/D10).
' Parsing and resolution are field-name agnostic. Clause model: <field> = <expr> joined by '|'.
' general.delimeter splits list values inside expressions only.
' B2: max-N offset form supported for scalar and range bound tokens (keyword_max_offset).
'     Offset grammar anchored to trimmed whole slice: max + optional ws + '-' + optional ws + [0-9]+.
'     min-N out of scope. No offsets on numeric literals. List elements are unchanged (literal tokens).
' T2/D10: min/max resolve from all available numeric values of the referenced field in the pivot
' (global per field), captured before applying data-file clauses so resolution is independent of other
' clauses' selections. Placeholder max_<field> resolution is handled separately in report-sheet flow
' (rev.7, D11) and is not derived from filter clause derivation/representatives.

Public Type ParsedClause
    fieldName As String      ' trimmed as configured
    fieldKey As String       ' normalized identity for duplicate detection / field-state lookup (Trim + LCase$)
    RawExpr As String        ' trimmed expression
    Kind As String           ' "scalar" | "range" | "list"
    ScalarValue As String    ' raw scalar token as trimmed (literal or keyword/max-offset token text)
    ScalarKind As String     ' "literal" | "min" | "max" | "max_offset" (scalar only)
    ScalarHasOffset As Boolean
    ScalarOffsetN As Long
    RangeStart As String     ' bound token raw text (min/max/max-N or numeric-convertible literal), trimmed
    RangeEnd As String       ' bound token raw text (min/max/max-N or numeric-convertible literal), trimmed
    RangeStartKind As String ' "min" | "max" | "max_offset" | "numeric"
    RangeEndKind As String   ' "min" | "max" | "max_offset" | "numeric"
    RangeStartHasOffset As Boolean
    RangeEndHasOffset As Boolean
    RangeStartOffsetN As Long
    RangeEndOffsetN As Long
    RangeStartIsNumeric As Boolean
    RangeEndIsNumeric As Boolean
    RangeStartNum As Double  ' valid when RangeStartIsNumeric = True
    RangeEndNum As Double    ' valid when RangeEndIsNumeric = True
    ListValues() As String
    HasList As Boolean
End Type

Public Type ParsedFilters
    Clauses() As ParsedClause
    Count As Long
End Type

' Generic resolved context for D10 filter bound resolution.
' FieldStates: Scripting.Dictionary keyed by FieldKey -> Variant array:
'   [0] HasNumericDomain (Boolean) - True when at least one numeric value is available for the field
'   [1] GlobalMin (Double)         - valid when HasNumericDomain
'   [2] GlobalMax (Double)         - valid when HasNumericDomain
'   [3] NumericDomain (Collection) - available numeric Double values (deduped), valid when HasNumericDomain
' Legacy placeholder aliases removed (rev.7): report_sheet max_<field> values are resolved independently.
Public Type ResolvedContext
    FieldStates As Object
End Type


' Uniform selection rule for all filters (T3 / Q2-G4,Q2-G5,Q2-G9), including report-level:
' - field missing/unavailable => fail
' - if none of specified values present/matchable => fail with clear contextual error (no skip-and-continue)
' - if at least one specified value present/matchable => apply matched subset without error

Public Function SplitReportValues(ByVal raw As String, ByVal delim As String, ByVal ctx As String) As Variant
    Dim arr() As String, i As Long, n As Long, outArr() As String
    If Len(delim) = 0 Then modTreportConfig.Fail 2309, ctx, "general.delimeter is required for multi-value report filters."
    arr = Split(raw, delim)
    n = -1
    For i = LBound(arr) To UBound(arr)
        arr(i) = modTreportConfig.Norm(arr(i))
        If Len(arr(i)) > 0 Then n = n + 1
    Next i
    If n < 0 Then modTreportConfig.Fail 2310, ctx, "No values provided for report filter."
    ReDim outArr(0 To n)
    n = -1
    For i = LBound(arr) To UBound(arr)
        If Len(arr(i)) > 0 Then
            n = n + 1
            outArr(n) = arr(i)
        End If
    Next i
    SplitReportValues = outArr
End Function

Private Function StripOuterParens(ByVal expr As String) As String
    Dim s As String
    s = modTreportConfig.Norm(expr)
    If Len(s) >= 2 Then
        If Left$(s, 1) = "(" And Right$(s, 1) = ")" Then
            StripOuterParens = Mid$(s, 2, Len(s) - 2)
            Exit Function
        End If
    End If
    StripOuterParens = s
End Function

Private Function IsNumericConvertibleToken(ByVal t As String, ByRef outNum As Double) As Boolean
    ' Deterministic numeric-convertible test for filter tokens/bounds (F4/Q2-G3).
    ' Accepts optional sign, digits with at most one decimal separator ('.' or ','),
    ' requiring at least one digit. Rejects empty, exponent, thousands-grouped, or junk forms.
    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim sepCount As Long
    Dim digitCount As Long
    Dim startPos As Long
    Dim normNum As String

    IsNumericConvertibleToken = False
    outNum = 0#
    s = modTreportConfig.Norm(t)
    If Len(s) = 0 Then Exit Function

    startPos = 1
    ch = Left$(s, 1)
    If ch = "+" Or ch = "-" Then startPos = 2
    If startPos > Len(s) Then Exit Function

    For i = startPos To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            digitCount = digitCount + 1
        ElseIf ch = "." Or ch = "," Then
            sepCount = sepCount + 1
            If sepCount > 1 Then Exit Function
        Else
            Exit Function
        End If
    Next i
    If digitCount = 0 Then Exit Function

    normNum = Replace$(s, ",", ".")
    On Error GoTo EH
    outNum = CDbl(normNum)
    IsNumericConvertibleToken = True
    Exit Function
EH:
    Err.Clear
    IsNumericConvertibleToken = False
    outNum = 0#
End Function

Private Function IsNumericConvertible(ByVal t As String) As Boolean
    Dim d As Double
    IsNumericConvertible = IsNumericConvertibleToken(t, d)
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then
            IsAllDigits = False
            Exit Function
        End If
    Next i
    IsAllDigits = True
End Function

Private Function TryParseMaxOffsetToken(ByVal t As String, ByRef outOffsetN As Long, ByRef outNormalized As String) As Boolean
    ' B2/RC-3: keyword_max_offset = "max" (case-insensitive) + optional ws + "-" + optional ws + digits [0-9]+
    ' consuming entire trimmed slice. Returns True only for complete offset form.
    Dim s As String
    Dim pDash As Long
    Dim leftPart As String
    Dim rightPart As String
    Dim nVal As Double

    TryParseMaxOffsetToken = False
    outOffsetN = 0
    outNormalized = vbNullString
    s = modTreportConfig.Norm(t)
    If Len(s) < 5 Then Exit Function ' "max-0"
    If LCase$(Left$(s, 3)) <> "max" Then Exit Function

    pDash = InStr(4, s, "-", vbBinaryCompare)
    If pDash = 0 Then Exit Function
    leftPart = modTreportConfig.Norm(Left$(s, pDash - 1))
    rightPart = modTreportConfig.Norm(Mid$(s, pDash + 1))
    If LCase$(leftPart) <> "max" Then Exit Function
    If Len(rightPart) = 0 Then Exit Function
    If Not IsAllDigits(rightPart) Then Exit Function

    On Error GoTo EHNum
    nVal = CDbl(rightPart)
    If nVal < 0# Or nVal > 2147483647# Then Exit Function
    outOffsetN = CLng(nVal)
    outNormalized = "max-" & rightPart
    TryParseMaxOffsetToken = True
    Exit Function
EHNum:
    Err.Clear
    TryParseMaxOffsetToken = False
End Function

Private Function IsMinMaxToken(ByVal t As String) As Boolean
    Dim s As String
    s = LCase$(modTreportConfig.Norm(t))
    IsMinMaxToken = (s = "min" Or s = "max")
End Function

Private Function IsBoundToken(ByVal t As String) As Boolean
    ' Range/scalar bound token: min, max, max-N offset (complete), or numeric-convertible literal.
    Dim d As Double
    Dim offN As Long
    Dim normOff As String
    If IsMinMaxToken(t) Then
        IsBoundToken = True
    ElseIf TryParseMaxOffsetToken(t, offN, normOff) Then
        IsBoundToken = True
    Else
        IsBoundToken = IsNumericConvertibleToken(t, d)
    End If
End Function

Private Sub ParseBoundToken(ByVal t As String, ByRef outKind As String, ByRef outHasOffset As Boolean, ByRef outOffsetN As Long, ByRef outIsNumeric As Boolean, ByRef outNum As Double)
    Dim s As String
    Dim normOff As String
    s = modTreportConfig.Norm(t)
    outKind = vbNullString
    outHasOffset = False
    outOffsetN = 0
    outIsNumeric = False
    outNum = 0#

    If TryParseMaxOffsetToken(s, outOffsetN, normOff) Then
        outKind = "max_offset"
        outHasOffset = True
        Exit Sub
    End If

    Select Case LCase$(s)
        Case "min"
            outKind = "min"
            Exit Sub
        Case "max"
            outKind = "max"
            Exit Sub
    End Select

    If IsNumericConvertibleToken(s, outNum) Then
        outKind = "numeric"
        outIsNumeric = True
        Exit Sub
    End If

    outKind = "invalid"
End Sub

Private Function CountValidRangeSplits(ByVal inner As String, ByRef outLeft As String, ByRef outRight As String) As Long
    Dim p As Long
    Dim leftB As String
    Dim rightB As String
    Dim foundCount As Long

    foundCount = 0
    outLeft = vbNullString
    outRight = vbNullString
    If Len(inner) < 3 Then
        CountValidRangeSplits = 0
        Exit Function
    End If

    For p = 2 To Len(inner) - 1
        If Mid$(inner, p, 1) = "-" Then
            leftB = modTreportConfig.Norm(Left$(inner, p - 1))
            rightB = modTreportConfig.Norm(Mid$(inner, p + 1))
            If IsBoundToken(leftB) And IsBoundToken(rightB) Then
                foundCount = foundCount + 1
                If foundCount = 1 Then
                    outLeft = leftB
                    outRight = rightB
                End If
            End If
        End If
    Next p
    CountValidRangeSplits = foundCount
End Function


Public Sub ClassifyExpr(ByRef c As ParsedClause, ByVal delim As String, ByVal ctx As String)
    ' Disambiguation (T1/B2): delimiter present => list (unchanged; max-N in list stays literal element token);
    ' else outer-paren with exactly one valid split into complete bound tokens => range; else scalar.
    ' RC-3/RC-4: complete bound tokens = numeric literal, min, max, max-N (entire trimmed slice). Offset '-'
    ' consumed inside bound token, never structural separator.
    Dim e As String
    Dim inner As String
    Dim parts() As String
    Dim i As Long
    Dim leftB As String
    Dim rightB As String
    Dim splitCount As Long
    Dim exprCtx As String

    e = modTreportConfig.Norm(c.RawExpr)
    c.RawExpr = e
    exprCtx = ctx & ",field='" & c.fieldName & "',expr='" & e & "'"

    If Len(e) = 0 Then modTreportConfig.Fail 2305, exprCtx, "Empty expression."
    If Len(delim) = 0 Then modTreportConfig.Fail 2306, ctx & ",field='" & c.fieldName & "'", "general.delimeter is required for filter parsing."

    c.Kind = vbNullString
    c.ScalarValue = vbNullString
    c.ScalarKind = vbNullString
    c.ScalarHasOffset = False
    c.ScalarOffsetN = 0
    c.RangeStart = vbNullString
    c.RangeEnd = vbNullString
    c.RangeStartKind = vbNullString
    c.RangeEndKind = vbNullString
    c.RangeStartHasOffset = False
    c.RangeEndHasOffset = False
    c.RangeStartOffsetN = 0
    c.RangeEndOffsetN = 0
    c.RangeStartIsNumeric = False
    c.RangeEndIsNumeric = False
    c.RangeStartNum = 0#
    c.RangeEndNum = 0#
    c.HasList = False

    ' List: expression contains general.delimeter (list values inside expressions only).
    ' RC-2: preserve existing list parsing/normalization/element handling exactly.
    If InStr(1, e, delim, vbBinaryCompare) > 0 Then
        c.Kind = "list"
        inner = StripOuterParens(e)
        parts = Split(inner, delim)
        ReDim c.ListValues(LBound(parts) To UBound(parts))
        For i = LBound(parts) To UBound(parts)
            c.ListValues(i) = modTreportConfig.Norm(parts(i))
            If Len(c.ListValues(i)) = 0 Then
                modTreportConfig.Fail 2311, exprCtx, "Empty value in list expression."
            End If
        Next i
        c.HasList = True
        Exit Sub
    End If

    ' Range: outer parentheses and exactly one valid split into two complete bound tokens.
    If Left$(e, 1) = "(" And Right$(e, 1) = ")" And Len(e) >= 5 Then
        inner = Mid$(e, 2, Len(e) - 2)
        splitCount = CountValidRangeSplits(inner, leftB, rightB)
        If splitCount = 1 Then
            c.Kind = "range"
            c.RangeStart = leftB
            c.RangeEnd = rightB
            ParseBoundToken leftB, c.RangeStartKind, c.RangeStartHasOffset, c.RangeStartOffsetN, c.RangeStartIsNumeric, c.RangeStartNum
            ParseBoundToken rightB, c.RangeEndKind, c.RangeEndHasOffset, c.RangeEndOffsetN, c.RangeEndIsNumeric, c.RangeEndNum

            ' Bounds must be min/max, max-N offset, or numeric-convertible literals.
            If Not (c.RangeStartKind = "min" Or c.RangeStartKind = "max" Or c.RangeStartKind = "max_offset" Or c.RangeStartIsNumeric) Then
                modTreportConfig.Fail 2322, exprCtx & ",token='" & leftB & "'", "Range start bound '" & leftB & "' is not a valid bound token (expected min, max, max-N, or numeric literal)."
            End If
            If Not (c.RangeEndKind = "min" Or c.RangeEndKind = "max" Or c.RangeEndKind = "max_offset" Or c.RangeEndIsNumeric) Then
                modTreportConfig.Fail 2323, exprCtx & ",token='" & rightB & "'", "Range end bound '" & rightB & "' is not a valid bound token (expected min, max, max-N, or numeric literal)."
            End If

            ' Parse-time bound validation when both bounds are numeric literals.
            If c.RangeStartIsNumeric And c.RangeEndIsNumeric Then
                If c.RangeStartNum > c.RangeEndNum Then
                    modTreportConfig.Fail 2319, exprCtx, "Range start is greater than range end after numeric validation."
                End If
            End If
            Exit Sub
        End If
    End If

    ' Scalar: complete max/min/max-N bound token => keyword semantics; else existing literal behavior.
    c.Kind = "scalar"
    c.ScalarValue = e
    Dim sk As String, sh As Boolean, sn As Long, sNum As Boolean, sNumV As Double
    ParseBoundToken e, sk, sh, sn, sNum, sNumV
    If sk = "min" Or sk = "max" Or sk = "max_offset" Then
        c.ScalarKind = sk
        c.ScalarHasOffset = sh
        c.ScalarOffsetN = sn
    Else
        c.ScalarKind = "literal"
        c.ScalarHasOffset = False
        c.ScalarOffsetN = 0
    End If
End Sub

' (Parsing helpers above provide deterministic bound/numeric checks for T1.)

Public Function ParseAndClassify(ByVal filtersText As String, ByVal delim As String, ByVal ctx As String) As ParsedFilters
    ' Generic parse/classify/validate for data-file/report filter expressions (T1).
    Dim pf As ParsedFilters
    Dim parts() As String
    Dim i As Long
    Dim n As Long
    Dim s As String
    Dim a As Long
    Dim b As Long
    Dim clause As String
    Dim eqPos As Long
    Dim fname As String
    Dim expr As String
    Dim clauseCtx As String

    s = modTreportConfig.Norm(filtersText)
    If Len(s) = 0 Then modTreportConfig.Fail 2301, ctx, "filters expression is empty."
    If Len(delim) = 0 Then modTreportConfig.Fail 2306, ctx, "general.delimeter is required for filter parsing."

    parts = Split(s, "|")
    n = UBound(parts) - LBound(parts) + 1
    If n <= 0 Then modTreportConfig.Fail 2302, ctx, "filters expression has no clauses."

    ReDim pf.Clauses(1 To n)
    pf.Count = n

    For i = LBound(parts) To UBound(parts)
        clause = modTreportConfig.Norm(parts(i))
        clauseCtx = ctx & ",clause[" & CStr(i - LBound(parts) + 1) & "]"
        If Len(clause) = 0 Then
            modTreportConfig.Fail 2302, clauseCtx & ",filters='" & s & "'", "Empty filter clause."
        End If

        eqPos = InStr(1, clause, "=", vbBinaryCompare)
        If eqPos <= 1 Then
            modTreportConfig.Fail 2303, clauseCtx & ",clause='" & clause & "'", "Invalid clause (expected '<field> = <expr>')."
        End If

        fname = modTreportConfig.Norm(Left$(clause, eqPos - 1))
        expr = modTreportConfig.Norm(Mid$(clause, eqPos + 1))
        If Len(fname) = 0 Or Len(expr) = 0 Then
            modTreportConfig.Fail 2304, clauseCtx & ",clause='" & clause & "'", "Invalid clause: field and expression must be non-empty."
        End If

        pf.Clauses(i - LBound(parts) + 1).fieldName = fname
        pf.Clauses(i - LBound(parts) + 1).fieldKey = LCase$(fname)
        pf.Clauses(i - LBound(parts) + 1).RawExpr = expr
        ClassifyExpr pf.Clauses(i - LBound(parts) + 1), delim, clauseCtx
    Next i

    ' Q2-G7: same field repeated in filters is not allowed (trimmed, case-insensitive duplicate detection).
    For a = 1 To pf.Count
        For b = a + 1 To pf.Count
            If StrComp(pf.Clauses(a).fieldKey, pf.Clauses(b).fieldKey, vbBinaryCompare) = 0 Then
                modTreportConfig.Fail 2312, ctx & ",field='" & pf.Clauses(a).fieldName & "'", "Duplicate filter field (case-insensitive): '" & pf.Clauses(a).fieldName & "' and '" & pf.Clauses(b).fieldName & "'."
            End If
        Next b
    Next a

    ParseAndClassify = pf
End Function

' --- T2/T4: generic per-field min/max resolution (global) and representative state ---

Private Function GetOrCreateFieldState(ByRef rc As ResolvedContext, ByVal pt As PivotTable, ByVal fieldName As String, ByVal fieldKey As String, ByVal ctx As String) As Variant
    ' Returns Variant array [0]=HasNumericDomain, [1]=GlobalMin, [2]=GlobalMax, [3]=NumericDomain(Collection).
    ' Captures all available numeric values of the referenced field from current pivot state (global per field).
    ' Must be called before applying data-file clauses (D10).
    Dim stateArr As Variant
    Dim pfld As PivotField
    Dim vals As Collection
    Dim hasNum As Boolean
    Dim minV As Double
    Dim maxV As Double

    If rc.FieldStates.Exists(fieldKey) Then
        GetOrCreateFieldState = rc.FieldStates(fieldKey)
        Exit Function
    End If

    Set pfld = modTreportPivot.GetPivotFieldRequired(pt, fieldName, ctx & ",field='" & fieldName & "'")
    Set vals = modTreportPivot.FieldAvailableNumericValues(pfld)
    If vals.Count > 0 Then
        hasNum = True
        minV = modTreportPivot.MinNumeric(vals)
        maxV = modTreportPivot.MaxNumeric(vals)
    Else
        hasNum = False
        minV = 0#
        maxV = 0#
    End If

    stateArr = Array(hasNum, minV, maxV, vals)
    rc.FieldStates.Add fieldKey, stateArr
    GetOrCreateFieldState = stateArr
End Function

Private Function BoundTokenCtx(ByVal token As String, ByVal kind As String, ByVal hasOffset As Boolean, ByVal offsetN As Long) As String
    Dim s As String
    s = ",token='" & token & "'"
    If hasOffset Or kind = "max_offset" Then s = s & ",N=" & CStr(offsetN)
    BoundTokenCtx = s
End Function

Private Function ResolveGlobalBound(ByVal token As String, ByVal boundKind As String, ByVal hasOffset As Boolean, ByVal offsetN As Long, ByVal numericValue As Double, ByVal isNumeric As Boolean, ByRef c As ParsedClause, ByRef rc As ResolvedContext, ByVal ctx As String) As Double
    ' Resolve min/max/max-N/numeric-literal bound token using pre-resolved global per-field state (T2/B2/D1/D2).
    Dim st As Variant
    Dim maxV As Double
    Dim computed As Double
    Dim domain As Collection
    Dim tokCtx As String
    Dim baseCtx As String

    If isNumeric Then
        ResolveGlobalBound = numericValue
        Exit Function
    End If

    tokCtx = BoundTokenCtx(token, boundKind, hasOffset, offsetN)
    baseCtx = ctx & ",field='" & c.fieldName & "',expr='" & c.RawExpr & "'" & tokCtx

    Select Case boundKind
        Case "min"
            st = rc.FieldStates(c.fieldKey)
            If Not CBool(st(0)) Then
                modTreportConfig.Fail 2313, baseCtx & ",resolved_max=<n/a>", "No numeric values available to resolve global '" & token & "'."
            End If
            ResolveGlobalBound = CDbl(st(1))
            Exit Function

        Case "max"
            st = rc.FieldStates(c.fieldKey)
            If Not CBool(st(0)) Then
                modTreportConfig.Fail 2313, baseCtx & ",resolved_max=<n/a>", "No numeric values available to resolve global '" & token & "'."
            End If
            ResolveGlobalBound = CDbl(st(2))
            Exit Function

        Case "max_offset"
            st = rc.FieldStates(c.fieldKey)
            If Not CBool(st(0)) Then
                modTreportConfig.Fail 2313, baseCtx & ",resolved_max=<n/a>,computed=<n/a>", "No numeric values/domain available to resolve offset bound '" & token & "'."
            End If
            maxV = CDbl(st(2))
            If CDbl(offsetN) < 0# Or CDbl(offsetN) >= maxV Then
                ' D1/RC-1: include field,resolved max,N,token/expr; computed included if derived/useful.
                computed = maxV - CDbl(offsetN)
                modTreportConfig.Fail 2324, baseCtx & ",resolved_max=" & CStr(maxV) & ",computed=" & CStr(computed), "Invalid max-offset: N must satisfy 0 <= N < max_value (field='" & c.fieldName & "', resolved_max=" & CStr(maxV) & ", N=" & CStr(offsetN) & ", computed=" & CStr(computed) & ", token='" & token & "')."
            End If
            computed = maxV - CDbl(offsetN)
            Set domain = st(3)
            If Not NumericDomainContains(domain, computed) Then
                ' D2/RC-1: include field,resolved max,N,computed and context.
                modTreportConfig.Fail 2325, baseCtx & ",resolved_max=" & CStr(maxV) & ",computed=" & CStr(computed), "Computed max-offset value is not present in field numeric domain (field='" & c.fieldName & "', resolved_max=" & CStr(maxV) & ", N=" & CStr(offsetN) & ", computed=" & CStr(computed) & ", token='" & token & "')."
            End If
            ResolveGlobalBound = computed
            Exit Function
    End Select

    modTreportConfig.Fail 2314, baseCtx, "Unsupported bound token '" & token & "'."
End Function

Private Function NumericDomainContains(ByVal domain As Collection, ByVal valueD As Double) As Boolean
    Dim i As Long
    If domain Is Nothing Then Exit Function
    For i = 1 To domain.Count
        If CDbl(domain(i)) = valueD Then
            NumericDomainContains = True
            Exit Function
        End If
    Next i
    NumericDomainContains = False
End Function

Public Function ResolveFilterContext(ByVal pt As PivotTable, ByRef pf As ParsedFilters, ByVal ctx As String) As ResolvedContext
    ' D10: generic per-field resolution. Captures global numeric domain per referenced field from pivot
    ' before any clause application, so min/max resolution is independent of other clauses' selections.
    ' rev.7: no placeholder aliases/representatives; report_sheet max_<field> resolved separately.
    Dim rc As ResolvedContext
    Dim i As Long

    Set rc.FieldStates = CreateObject("Scripting.Dictionary")
    rc.FieldStates.CompareMode = vbBinaryCompare

    ' Pre-resolve global numeric domains for all referenced fields (once per unique field).
    For i = 1 To pf.Count
        GetOrCreateFieldState rc, pt, pf.Clauses(i).fieldName, pf.Clauses(i).fieldKey, ctx
    Next i

    ResolveFilterContext = rc
End Function

Public Sub ApplyDataFileFilters(ByVal pt As PivotTable, ByRef pf As ParsedFilters, ByVal delim As String, ByRef rc As ResolvedContext, ByVal ctx As String)
    ' Deterministic AND in declared clause order. Desired values use global per-field min/max from rc
    ' (resolved before any clause application) so each field is independent of other clauses' selections (D10).
    Dim i As Long
    For i = 1 To pf.Count
        ApplySingleClauseSelection pt, pf.Clauses(i), delim, rc, ctx & ",source=data_files.filters"
    Next i
    On Error Resume Next
    pt.Update
    DoEvents
    On Error GoTo 0
End Sub

Private Sub ApplySingleClauseSelection(ByVal pt As PivotTable, ByRef c As ParsedClause, ByVal delim As String, ByRef rc As ResolvedContext, ByVal ctx As String)
    Dim desired As Collection
    Dim clauseCtx As String
    clauseCtx = ctx & ",field='" & c.fieldName & "',expr='" & c.RawExpr & "'"
    Set desired = BuildDesiredValues(pt, c, rc, clauseCtx)
    ' Uniform rule (D9): apply matched subset if >=1 value matches; fail if none match.
    modTreportPivot.ApplyFieldSelection pt, c.fieldName, desired, clauseCtx
End Sub

Private Function BuildDesiredValues(ByVal pt As PivotTable, ByRef c As ParsedClause, ByRef rc As ResolvedContext, ByVal ctx As String) As Collection
    ' Build desired pivot-item values for a clause using generic global per-field min/max/max-N resolution (D10/B2/D1/D2).
    ' No placeholder representative tracking (rev.7: max_<field> resolved independently in report-sheet flow).
    Dim col As New Collection
    Dim pfld As PivotField
    Dim domain As Collection
    Dim st As Variant
    Dim loB As Double, hiB As Double
    Dim i As Long, j As Long
    Dim matched As Boolean
    Dim v As Variant
    Dim resolvedNum As Double

    Set pfld = modTreportPivot.GetPivotFieldRequired(pt, c.fieldName, ctx)
    Select Case c.Kind
        Case "scalar"
            If c.ScalarKind = "min" Or c.ScalarKind = "max" Or c.ScalarKind = "max_offset" Then
                resolvedNum = ResolveGlobalBound(c.ScalarValue, c.ScalarKind, c.ScalarHasOffset, c.ScalarOffsetN, 0#, False, c, rc, ctx)
                col.Add NormalizeToPivotItem(pfld, CStr(resolvedNum))
            Else
                col.Add NormalizeToPivotItem(pfld, modTreportConfig.Norm(c.ScalarValue))
            End If
        Case "range"
            loB = ResolveGlobalBound(c.RangeStart, c.RangeStartKind, c.RangeStartHasOffset, c.RangeStartOffsetN, c.RangeStartNum, c.RangeStartIsNumeric, c, rc, ctx & ",range_start")
            hiB = ResolveGlobalBound(c.RangeEnd, c.RangeEndKind, c.RangeEndHasOffset, c.RangeEndOffsetN, c.RangeEndNum, c.RangeEndIsNumeric, c, rc, ctx & ",range_end")
            If loB > hiB Then modTreportConfig.Fail 2319, ctx & ",field='" & c.fieldName & "',expr='" & c.RawExpr & "'", "Range start is greater than range end after resolving bounds (" & CStr(loB) & " > " & CStr(hiB) & ")."
            st = rc.FieldStates(c.fieldKey)
            If Not CBool(st(0)) Then
                modTreportConfig.Fail 2320, ctx & ",field='" & c.fieldName & "',expr='" & c.RawExpr & "'", "No available numeric values matched range (" & c.RangeStart & "-" & c.RangeEnd & ")."
            End If
            Set domain = st(3)
            matched = False
            For i = 1 To domain.Count
                v = domain(i)
                If CDbl(v) >= loB And CDbl(v) <= hiB Then
                    col.Add NormalizeToPivotItem(pfld, CStr(v))
                    matched = True
                End If
            Next i
            If Not matched Then modTreportConfig.Fail 2320, ctx & ",field='" & c.fieldName & "',expr='" & c.RawExpr & "'", "No available numeric values matched range (" & c.RangeStart & "-" & c.RangeEnd & ")."
        Case "list"
            For j = LBound(c.ListValues) To UBound(c.ListValues)
                col.Add NormalizeToPivotItem(pfld, c.ListValues(j))
            Next j
        Case Else
            modTreportConfig.Fail 2321, ctx & ",field='" & c.fieldName & "',expr='" & c.RawExpr & "'", "Unsupported filter expression kind."
    End Select
    Set BuildDesiredValues = col
End Function

Public Function NormalizeToPivotItem(ByVal pf As PivotField, ByVal token As String) As String
    Dim t As String, i As Long, n As String
    t = modTreportConfig.Norm(token)
    On Error Resume Next
    For i = 1 To pf.PivotItems.Count
        n = modTreportConfig.Norm(pf.PivotItems(i).Name)
        If StrComp(n, t, vbBinaryCompare) = 0 Then
            NormalizeToPivotItem = n
            Exit Function
        End If
        If IsNumeric(n) And IsNumeric(t) Then
            If CDbl(n) = CDbl(t) Then
                NormalizeToPivotItem = n
                Exit Function
            End If
        End If
    Next i
    On Error GoTo 0
    NormalizeToPivotItem = t
End Function

Public Sub ApplyReportFilters(ByVal pt As PivotTable, ByRef rep As modTreportConfig.TreportReport, ByVal delim As String, ByVal ctx As String)
    Dim arr As Variant
    Dim colCat As New Collection
    Dim colTa As New Collection
    Dim i As Long
    Dim catCtx As String
    Dim taCtx As String

    ' Report-level filters use the same uniform selection rule as data-file filters (Q2-G9, D9):
    ' field missing/unavailable => fail; none of specified values present/matchable => fail;
    ' at least one present/matchable => apply matched subset without error (no skip-and-continue).
    catCtx = ctx & ",source=reports.articleLevel4Name_clean,field='articleLevel4Name_clean',values='" & modTreportConfig.Norm(rep.ArticleLevel4NameClean) & "'"
    arr = SplitReportValues(rep.ArticleLevel4NameClean, delim, catCtx)
    For i = LBound(arr) To UBound(arr)
        colCat.Add CStr(arr(i))
    Next i
    modTreportPivot.ApplyFieldSelection pt, "articleLevel4Name_clean", colCat, catCtx

    taCtx = ctx & ",source=reports.targetAudience,field='targetAudience',values='" & modTreportConfig.Norm(rep.TargetAudience) & "'"
    arr = SplitReportValues(rep.TargetAudience, delim, taCtx)
    For i = LBound(arr) To UBound(arr)
        colTa.Add CStr(arr(i))
    Next i
    modTreportPivot.ApplyFieldSelection pt, "targetAudience", colTa, taCtx
End Sub
