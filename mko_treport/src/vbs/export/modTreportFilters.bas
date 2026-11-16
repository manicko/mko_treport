Attribute VB_Name = "modTreportFilters"
Option Explicit

' Generic filter parsing/classification/validation (T1, spec 002 / Q2-G1..G7).
' Parsing is field-name agnostic: no hardcoded field semantics in generic parsing.
' Clause model: <field> = <expr> joined by '|'. general.delimeter splits list values inside expressions only.

Public Type ParsedClause
    fieldName As String      ' trimmed as configured
    FieldKey As String       ' normalized identity for duplicate detection (Trim + LCase$)
    RawExpr As String        ' trimmed expression
    Kind As String           ' "scalar" | "range" | "list"
    ScalarValue As String
    RangeStart As String     ' bound token (min/max or numeric-convertible literal), trimmed
    RangeEnd As String       ' bound token (min/max or numeric-convertible literal), trimmed
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

' Legacy placeholder/resolution context kept for current call-site compatibility (T2-T5 will evolve).
Public Type ResolvedContext
    YearFieldName As String
    MaxYear As Variant
    MaxMonth As Variant
    MaxWeek As Variant
    HasYear As Boolean
    HasMonth As Boolean
    HasWeek As Boolean
End Type


' Generalized articleLevel4Name_clean multi-value rule:
' - for report-level articleLevel4Name_clean selections, any configured token not found
'   among available pivot items is ignored (matched subset is applied)
' - if none of the configured tokens match, the caller must skip that report with warning
'   (error 2411) and continue other reports; this is not a whole-run failure.
' targetAudience and data-file filters remain strict (missing required values fail-fast).

Public Function SplitReportValues(ByVal raw As String, ByVal delim As String) As Variant
    Dim arr() As String, i As Long, n As Long, outArr() As String
    If Len(delim) = 0 Then modTreportConfig.Fail 2309, "reports", "general.delimeter is required for multi-value report filters."
    arr = Split(raw, delim)
    n = -1
    For i = LBound(arr) To UBound(arr)
        arr(i) = modTreportConfig.Norm(arr(i))
        If Len(arr(i)) > 0 Then n = n + 1
    Next i
    If n < 0 Then modTreportConfig.Fail 2310, "reports", "No values provided for report filter."
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

Private Function IsMinMaxToken(ByVal t As String) As Boolean
    Dim s As String
    s = LCase$(modTreportConfig.Norm(t))
    IsMinMaxToken = (s = "min" Or s = "max")
End Function

Private Function IsBoundToken(ByVal t As String) As Boolean
    ' Range bound token: min/max (case-insensitive) or numeric-convertible literal (Q2-G2/Q2-G3).
    Dim d As Double
    IsBoundToken = IsMinMaxToken(t) Or IsNumericConvertibleToken(t, d)
End Function

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
    ' Disambiguation (T1): delimiter present => list; else single valid outer-paren range => range; else scalar.
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
    c.RangeStart = vbNullString
    c.RangeEnd = vbNullString
    c.RangeStartIsNumeric = False
    c.RangeEndIsNumeric = False
    c.RangeStartNum = 0#
    c.RangeEndNum = 0#
    c.HasList = False

    ' List: expression contains general.delimeter (list values inside expressions only).
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

    ' Range: outer parentheses and exactly one valid split into two bound tokens.
    If Left$(e, 1) = "(" And Right$(e, 1) = ")" And Len(e) >= 5 Then
        inner = Mid$(e, 2, Len(e) - 2)
        splitCount = CountValidRangeSplits(inner, leftB, rightB)
        If splitCount = 1 Then
            c.Kind = "range"
            c.RangeStart = leftB
            c.RangeEnd = rightB
            c.RangeStartIsNumeric = IsNumericConvertibleToken(leftB, c.RangeStartNum)
            c.RangeEndIsNumeric = IsNumericConvertibleToken(rightB, c.RangeEndNum)

            ' Q2-G3: bounds must be min/max or numeric-convertible literals (range for non-numeric values invalid).
            If Not (IsMinMaxToken(leftB) Or c.RangeStartIsNumeric) Then
                modTreportConfig.Fail 2322, exprCtx, "Range start bound '" & leftB & "' is not numeric-convertible and not min/max."
            End If
            If Not (IsMinMaxToken(rightB) Or c.RangeEndIsNumeric) Then
                modTreportConfig.Fail 2323, exprCtx, "Range end bound '" & rightB & "' is not numeric-convertible and not min/max."
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

    c.Kind = "scalar"
    c.ScalarValue = e
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
        pf.Clauses(i - LBound(parts) + 1).FieldKey = LCase$(fname)
        pf.Clauses(i - LBound(parts) + 1).RawExpr = expr
        ClassifyExpr pf.Clauses(i - LBound(parts) + 1), delim, clauseCtx
    Next i

    ' Q2-G7: same field repeated in filters is not allowed (trimmed, case-insensitive duplicate detection).
    For a = 1 To pf.Count
        For b = a + 1 To pf.Count
            If StrComp(pf.Clauses(a).FieldKey, pf.Clauses(b).FieldKey, vbBinaryCompare) = 0 Then
                modTreportConfig.Fail 2312, ctx & ",field='" & pf.Clauses(a).fieldName & "'", "Duplicate filter field (case-insensitive): '" & pf.Clauses(a).fieldName & "' and '" & pf.Clauses(b).fieldName & "'."
            End If
        Next b
    Next a

    ParseAndClassify = pf
End Function

Private Function FindClauseIndex(ByRef pf As ParsedFilters, ByVal fieldName As String) As Long
    Dim i As Long
    For i = 1 To pf.Count
        If StrComp(pf.Clauses(i).fieldName, fieldName, vbBinaryCompare) = 0 Then
            FindClauseIndex = i
            Exit Function
        End If
    Next i
    FindClauseIndex = 0
End Function

Private Function ResolveBound(ByVal token As String, ByVal fieldName As String, ByVal pt As PivotTable, ByVal ctx As String) As Double
    Dim t As String
    Dim pfld As PivotField
    Dim vals As Collection
    t = LCase$(modTreportConfig.Norm(token))
    If IsNumeric(t) Then
        ResolveBound = CDbl(t)
        Exit Function
    End If
    Set pfld = modTreportPivot.GetPivotFieldRequired(pt, fieldName, ctx)
    Set vals = modTreportPivot.FieldAvailableNumericValues(pfld)
    If vals.Count = 0 Then modTreportConfig.Fail 2313, ctx & ",field='" & fieldName & "'", "No numeric values available to resolve '" & token & "'."
    If t = "min" Then
        ResolveBound = modTreportPivot.MinNumeric(vals)
    ElseIf t = "max" Then
        ResolveBound = modTreportPivot.MaxNumeric(vals)
    Else
        modTreportConfig.Fail 2314, ctx & ",field='" & fieldName & "'", "Unsupported bound token '" & token & "'."
    End If
End Function

Public Function ResolveMinMaxAndYear(ByVal pt As PivotTable, ByRef pf As ParsedFilters, ByVal ctx As String) As ResolvedContext
    Dim rc As ResolvedContext
    Dim yi As Long, mi As Long, wi As Long
    yi = FindClauseIndex(pf, "Year")
    mi = FindClauseIndex(pf, "Month Number")
    wi = FindClauseIndex(pf, "Week Number")

    If yi > 0 Then
        rc.YearFieldName = "Year"
        rc.HasYear = True
        Select Case pf.Clauses(yi).Kind
            Case "scalar"
                Dim ytok As String
                ytok = LCase$(modTreportConfig.Norm(pf.Clauses(yi).ScalarValue))
                If ytok = "max" Or ytok = "min" Then
                    rc.MaxYear = ResolveBound(ytok, "Year", pt, ctx & ",resolve_year")
                ElseIf IsNumeric(ytok) Then
                    rc.MaxYear = CDbl(ytok)
                Else
                    modTreportConfig.Fail 2315, ctx & ",field='Year'", "Unsupported scalar year value '" & pf.Clauses(yi).ScalarValue & "'."
                End If
            Case "range"
                rc.MaxYear = ResolveBound(pf.Clauses(yi).RangeEnd, "Year", pt, ctx & ",resolve_year_end")
            Case "list"
                rc.MaxYear = MaxOfStringArray(pf.Clauses(yi).ListValues, ctx & ",field='Year'")
        End Select
        ApplySingleClauseSelection pt, pf.Clauses(yi), vbNullString, rc, ctx & ",apply_year_context"
        On Error Resume Next
        pt.Update
        DoEvents
        On Error GoTo 0
    End If

    If mi > 0 Then
        rc.HasMonth = True
        rc.MaxMonth = ResolveClauseMax(pt, pf.Clauses(mi), "Month Number", ctx)
    End If
    If wi > 0 Then
        rc.HasWeek = True
        rc.MaxWeek = ResolveClauseMax(pt, pf.Clauses(wi), "Week Number", ctx)
    End If

    ResolveMinMaxAndYear = rc
End Function

Private Function MaxOfStringArray(ByRef arr() As String, ByVal ctx As String) As Double
    Dim i As Long, hasAny As Boolean, v As Double
    hasAny = False
    For i = LBound(arr) To UBound(arr)
        If IsNumeric(modTreportConfig.Norm(arr(i))) Then
            v = CDbl(modTreportConfig.Norm(arr(i)))
            If Not hasAny Or v > MaxOfStringArray Then MaxOfStringArray = v: hasAny = True
        End If
    Next i
    If Not hasAny Then modTreportConfig.Fail 2316, ctx, "No numeric values in list to derive maximum."
End Function

Private Function ResolveClauseMax(ByVal pt As PivotTable, ByRef c As ParsedClause, ByVal defaultField As String, ByVal ctx As String) As Variant
    Select Case c.Kind
        Case "scalar"
            Dim t As String
            t = LCase$(modTreportConfig.Norm(c.ScalarValue))
            If t = "max" Or t = "min" Then
                ResolveClauseMax = ResolveBound(t, c.fieldName, pt, ctx & ",field='" & c.fieldName & "'")
            ElseIf IsNumeric(t) Then
                ResolveClauseMax = CDbl(t)
            Else
                modTreportConfig.Fail 2317, ctx & ",field='" & c.fieldName & "'", "Unsupported scalar value '" & c.ScalarValue & "'."
            End If
        Case "range"
            ResolveClauseMax = ResolveBound(c.RangeEnd, c.fieldName, pt, ctx & ",field='" & c.fieldName & "',range_end")
        Case "list"
            ResolveClauseMax = MaxOfStringArray(c.ListValues, ctx & ",field='" & c.fieldName & "'")
        Case Else
            modTreportConfig.Fail 2318, ctx & ",field='" & c.fieldName & "'", "Unsupported expression kind."
    End Select
End Function

Public Sub ApplyDataFileFilters(ByVal pt As PivotTable, ByRef pf As ParsedFilters, ByVal delim As String, ByRef rc As ResolvedContext, ByVal ctx As String)
    Dim i As Long
    For i = 1 To pf.Count
        ApplySingleClauseSelection pt, pf.Clauses(i), delim, rc, ctx & ",field='" & pf.Clauses(i).fieldName & "'"
    Next i
    On Error Resume Next
    pt.Update
    DoEvents
    On Error GoTo 0
End Sub

Private Sub ApplySingleClauseSelection(ByVal pt As PivotTable, ByRef c As ParsedClause, ByVal delim As String, ByRef rc As ResolvedContext, ByVal ctx As String)
    Dim desired As Collection
    Set desired = BuildDesiredValues(pt, c, rc, ctx)
    ' data_files.filters remain strict: missing/unmatchable required values fail-fast.
    modTreportPivot.ApplyFieldSelection pt, c.fieldName, desired, False, False, ctx
End Sub

Private Function BuildDesiredValues(ByVal pt As PivotTable, ByRef c As ParsedClause, ByRef rc As ResolvedContext, ByVal ctx As String) As Collection
    Dim col As New Collection
    Dim pfld As PivotField
    Dim vals As Collection
    Dim v As Variant
    Dim loB As Double, hiB As Double
    Dim s As String, i As Long, j As Long, matched As Boolean

    Set pfld = modTreportPivot.GetPivotFieldRequired(pt, c.fieldName, ctx)
    Select Case c.Kind
        Case "scalar"
            s = modTreportConfig.Norm(c.ScalarValue)
            If LCase$(s) = "max" Or LCase$(s) = "min" Then
                s = CStr(ResolveBound(s, c.fieldName, pt, ctx))
            End If
            col.Add NormalizeToPivotItem(pfld, s)
        Case "range"
            loB = ResolveBound(c.RangeStart, c.fieldName, pt, ctx)
            hiB = ResolveBound(c.RangeEnd, c.fieldName, pt, ctx)
            If loB > hiB Then modTreportConfig.Fail 2319, ctx & ",field='" & c.fieldName & "'", "Range start is greater than range end."
            Set vals = modTreportPivot.FieldAvailableNumericValues(pfld)
            matched = False
            For i = 1 To vals.Count
                v = vals(i)
                If CDbl(v) >= loB And CDbl(v) <= hiB Then
                    col.Add NormalizeToPivotItem(pfld, CStr(v))
                    matched = True
                End If
            Next i
            If Not matched Then modTreportConfig.Fail 2320, ctx & ",field='" & c.fieldName & "'", "No available numeric values matched range (" & c.RangeStart & "-" & c.RangeEnd & ")."
        Case "list"
            For j = LBound(c.ListValues) To UBound(c.ListValues)
                col.Add NormalizeToPivotItem(pfld, c.ListValues(j))
            Next j
        Case Else
            modTreportConfig.Fail 2321, ctx & ",field='" & c.fieldName & "'", "Unsupported filter expression kind."
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

    ' articleLevel4Name_clean (report-level, multi-value): generalized rule -
    ' apply matched subset of configured tokens; if none match, signal report-skip (2411).
    arr = SplitReportValues(rep.ArticleLevel4NameClean, delim)
    For i = LBound(arr) To UBound(arr)
        colCat.Add CStr(arr(i))
    Next i
    modTreportPivot.ApplyFieldSelection pt, "articleLevel4Name_clean", colCat, True, True, ctx & ",field='articleLevel4Name_clean'"

    ' targetAudience behavior unchanged (strict).
    arr = SplitReportValues(rep.TargetAudience, delim)
    For i = LBound(arr) To UBound(arr)
        colTa.Add CStr(arr(i))
    Next i
    modTreportPivot.ApplyFieldSelection pt, "targetAudience", colTa, False, False, ctx & ",field='targetAudience'"
End Sub
