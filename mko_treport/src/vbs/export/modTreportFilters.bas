Attribute VB_Name = "modTreportFilters"
Option Explicit

Public Type ParsedClause
    fieldName As String
    RawExpr As String
    Kind As String
    ScalarValue As String
    RangeStart As String
    RangeEnd As String
    ListValues() As String
    HasList As Boolean
End Type

Public Type ParsedFilters
    Clauses() As ParsedClause
    Count As Long
End Type

Public Type ResolvedContext
    YearFieldName As String
    MaxYear As Variant
    MaxMonth As Variant
    MaxWeek As Variant
    HasYear As Boolean
    HasMonth As Boolean
    HasWeek As Boolean
End Type


Public Function IsOptionalSkipToken(ByVal token As String) As Boolean
    Dim t As String
    t = modTreportConfig.Norm(token)
    ' V-008: "ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ" optional if absent
    If Len(t) = 0 Then Exit Function
    If StrComp(t, ChrW$(1047) & ChrW$(1059) & ChrW$(1041) & ChrW$(1053) & ChrW$(1040) & ChrW$(1071) & " " & _
               ChrW$(1055) & ChrW$(1040) & ChrW$(1057) & ChrW$(1058) & ChrW$(1040) & " " & _
               ChrW$(1044) & ChrW$(1051) & ChrW$(1071) & " " & _
               ChrW$(1044) & ChrW$(1045) & ChrW$(1058) & ChrW$(1045) & ChrW$(1049), vbBinaryCompare) = 0 Then
        IsOptionalSkipToken = True
        Exit Function
    End If
    ' robust fallback marker
    If InStr(1, t, ChrW$(1044) & ChrW$(1051) & ChrW$(1071) & " " & ChrW$(1044) & ChrW$(1045) & ChrW$(1058) & ChrW$(1045) & ChrW$(1049), vbBinaryCompare) > 0 Then
        If InStr(1, t, ChrW$(1047) & ChrW$(1059) & ChrW$(1041) & ChrW$(1053) & ChrW$(1040) & ChrW$(1071), vbBinaryCompare) > 0 Then
            IsOptionalSkipToken = True
        End If
    End If
End Function

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


Public Sub ClassifyExpr(ByRef c As ParsedClause, ByVal delim As String, ByVal ctx As String)
    Dim e As String, inner As String, parts() As String, i As Long
    Dim leftB As String, rightB As String, p As Long, found As Boolean
    e = modTreportConfig.Norm(c.RawExpr)
    If Len(e) = 0 Then modTreportConfig.Fail 2305, ctx & ",field='" & c.fieldName & "'", "Empty expression."
    If Len(delim) = 0 Then modTreportConfig.Fail 2306, ctx, "general.delimeter is required for filter parsing."
    If InStr(1, e, delim, vbBinaryCompare) > 0 Then
        c.Kind = "list"
        inner = StripOuterParens(e)
        parts = Split(inner, delim)
        ReDim c.ListValues(LBound(parts) To UBound(parts))
        For i = LBound(parts) To UBound(parts)
            c.ListValues(i) = modTreportConfig.Norm(parts(i))
            If Len(c.ListValues(i)) = 0 Then modTreportConfig.Fail 2311, ctx & ",field='" & c.fieldName & "'", "Empty value in list expression."
        Next i
        c.HasList = True
        Exit Sub
    End If
    If Left$(e, 1) = "(" And Right$(e, 1) = ")" And Len(e) >= 5 Then
        inner = Mid$(e, 2, Len(e) - 2)
        found = False
        For p = 2 To Len(inner) - 1
            If Mid$(inner, p, 1) = "-" Then
                leftB = modTreportConfig.Norm(Left$(inner, p - 1))
                rightB = modTreportConfig.Norm(Mid$(inner, p + 1))
                If IsBoundToken(leftB) And IsBoundToken(rightB) Then
                    c.Kind = "range"
                    c.RangeStart = leftB
                    c.RangeEnd = rightB
                    found = True
                    Exit For
                End If
            End If
        Next p
        If found Then Exit Sub
    End If
    c.Kind = "scalar"
    c.ScalarValue = e
End Sub

Private Function IsBoundToken(ByVal t As String) As Boolean
    Dim s As String
    s = LCase$(modTreportConfig.Norm(t))
    IsBoundToken = (s = "min" Or s = "max" Or IsNumeric(s))
End Function

Public Function ParseAndClassify(ByVal filtersText As String, ByVal delim As String, ByVal ctx As String) As ParsedFilters
    Dim pf As ParsedFilters
    Dim parts() As String, i As Long, n As Long
    Dim s As String, a As Long, b As Long
    s = modTreportConfig.Norm(filtersText)
    If Len(s) = 0 Then modTreportConfig.Fail 2301, ctx, "filters expression is empty."
    parts = Split(s, "|")
    n = 0
    For i = LBound(parts) To UBound(parts)
        If Len(modTreportConfig.Norm(parts(i))) > 0 Then n = n + 1
    Next i
    If n = 0 Then modTreportConfig.Fail 2302, ctx, "filters expression has no clauses."
    ReDim pf.Clauses(1 To n)
    pf.Count = n
    n = 0
    For i = LBound(parts) To UBound(parts)
        Dim clause As String, eqPos As Long, fname As String, expr As String
        clause = modTreportConfig.Norm(parts(i))
        If Len(clause) = 0 Then GoTo NextClause
        eqPos = InStr(1, clause, "=", vbBinaryCompare)
        If eqPos <= 1 Then modTreportConfig.Fail 2303, ctx, "Invalid clause (expected '<field> = <expr>'): '" & clause & "'"
        fname = modTreportConfig.Norm(Left$(clause, eqPos - 1))
        expr = modTreportConfig.Norm(Mid$(clause, eqPos + 1))
        If Len(fname) = 0 Or Len(expr) = 0 Then modTreportConfig.Fail 2304, ctx, "Invalid clause field/expression: '" & clause & "'"
        n = n + 1
        pf.Clauses(n).fieldName = fname
        pf.Clauses(n).RawExpr = expr
        ClassifyExpr pf.Clauses(n), delim, ctx
NextClause:
    Next i
    For a = 1 To pf.Count
        For b = a + 1 To pf.Count
            If StrComp(pf.Clauses(a).fieldName, pf.Clauses(b).fieldName, vbBinaryCompare) = 0 Then
                modTreportConfig.Fail 2312, ctx, "Duplicate filter field: '" & pf.Clauses(a).fieldName & "'."
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
    modTreportPivot.ApplyFieldSelection pt, c.fieldName, desired, False, ctx
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

    arr = SplitReportValues(rep.ArticleLevel4NameClean, delim)
    For i = LBound(arr) To UBound(arr)
        colCat.Add CStr(arr(i))
    Next i
    modTreportPivot.ApplyFieldSelection pt, "articleLevel4Name_clean", colCat, True, ctx & ",field='articleLevel4Name_clean'"

    arr = SplitReportValues(rep.TargetAudience, delim)
    For i = LBound(arr) To UBound(arr)
        colTa.Add CStr(arr(i))
    Next i
    modTreportPivot.ApplyFieldSelection pt, "targetAudience", colTa, False, ctx & ",field='targetAudience'"
End Sub
