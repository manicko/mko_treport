Attribute VB_Name = "modTreportPivot"
Option Explicit

' B3/D11 (rev.7): max_<field> and max-N_<field> extraction enumerates named field available items via
' PivotTable.PivotFields(fieldName).PivotItems in the row pivot/data context.
' Numeric-convertible detection is deterministic (Trim + sign/digit/decimal-sep parse) and
' independent of data_files.filters application/selection/representatives.
' D1/D2 (B3): N>=0, 0<=N<max_value, N==0==max; N>=max_value or computed value absent from
' pre-filter numeric-convertible field item domain (exact CDbl equality, PivotItems+TryParseNumericConvertible
' parity) => fail fast by callers with clear context.

Public Function GetDataSheet(ByVal wb As Workbook, ByVal sheetName As String, ByVal ctx As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then modTreportConfig.Fail 2401, ctx, "Data sheet '" & sheetName & "' was not found."
    Set GetDataSheet = ws
End Function

Public Function GetPrimaryPivot(ByVal ws As Worksheet, ByVal ctx As String) As PivotTable
    If ws.PivotTables.Count = 0 Then
        modTreportConfig.Fail 2402, ctx, "No PivotTable found on sheet '" & ws.Name & "'."
    End If
    Dim pt As PivotTable
    On Error Resume Next
    Set pt = ws.PivotTables("PivotTable1")
    On Error GoTo 0
    If pt Is Nothing Then Set pt = ws.PivotTables(1)
    Set GetPrimaryPivot = pt
End Function

Public Function TryGetPivotField(ByVal pt As PivotTable, ByVal fieldName As String) As PivotField
    Dim pf As PivotField
    On Error Resume Next
    Set pf = pt.PivotFields(fieldName)
    On Error GoTo 0
    Set TryGetPivotField = pf
End Function

Public Function GetPivotFieldRequired(ByVal pt As PivotTable, ByVal fieldName As String, ByVal ctx As String) As PivotField
    ' P01/D7/D12: availability is presence in PivotTable available fields (pt.PivotFields).
    ' A found-but-hidden/unplaced field is available; missing from PivotFields is unavailable.
    Dim pf As PivotField
    Set pf = TryGetPivotField(pt, fieldName)
    If pf Is Nothing Then
        modTreportConfig.Fail 2403, ctx & ",field='" & fieldName & "'", "Required pivot field '" & fieldName & "' was not found or unavailable in pivot context (D9: missing field in PivotTable available fields)."
    End If
    Set GetPivotFieldRequired = pf
End Function

Private Function TryParseNumericConvertible(ByVal raw As String, ByRef outNum As Double) As Boolean
    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim sepCount As Long
    Dim digitCount As Long
    Dim startPos As Long
    Dim normNum As String

    TryParseNumericConvertible = False
    outNum = 0#
    s = modTreportConfig.Norm(raw)
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
    TryParseNumericConvertible = True
    Exit Function
EH:
    Err.Clear
    TryParseNumericConvertible = False
    outNum = 0#
End Function

Private Sub CollectPivotFieldNumericDomain(ByVal pt As PivotTable, ByVal fieldName As String, ByVal ctx As String, ByRef outHasNum As Boolean, ByRef outMaxVal As Double, ByRef outDomain As Object)
    ' Single pre-filter enumeration of PivotItems.Name with TryParseNumericConvertible parity.
    ' outDomain keys are CStr(CDbl(numVal)) for exact numeric membership checks (CDbl equality semantics).
    Dim pf As PivotField
    Dim i As Long
    Dim itemName As String
    Dim numVal As Double
    Dim key As String

    Set pf = GetPivotFieldRequired(pt, fieldName, ctx)
    Set outDomain = CreateObject("Scripting.Dictionary")
    outDomain.CompareMode = vbBinaryCompare
    outHasNum = False
    outMaxVal = 0#

    On Error Resume Next
    For i = 1 To pf.PivotItems.Count
        itemName = CStr(pf.PivotItems(i).Name)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextItem
        End If
        If TryParseNumericConvertible(itemName, numVal) Then
            If Not outHasNum Or numVal > outMaxVal Then
                outMaxVal = numVal
                outHasNum = True
            End If
            key = CStr(CDbl(numVal))
            If Not outDomain.Exists(key) Then outDomain.Add key, True
        End If
NextItem:
        Err.Clear
    Next i
    On Error GoTo 0

    If Not outHasNum Then
        modTreportConfig.Fail 2619, ctx & ",field='" & fieldName & "'", "Cannot determine max_<" & fieldName & ">: no numeric-convertible values available for referenced PivotTable field in pivot context."
    End If
End Sub

Public Function GetPivotFieldNumericMaxValue(ByVal pt As PivotTable, ByVal fieldName As String, ByVal ctx As String) As Variant
    ' Deterministic maximum numeric value over all available numeric-convertible items of the named
    ' PivotTable field in the relevant pivot/data context for the processed data_files row.
    ' Enumeration source: PivotField.PivotItems(item).Name from PivotFields(fieldName) (pre-filter,
    ' independent of filter clause derivation/representatives and selected-subset results).
    Dim hasNum As Boolean
    Dim maxVal As Double
    Dim domain As Object

    CollectPivotFieldNumericDomain pt, fieldName, ctx, hasNum, maxVal, domain

    ' Return numeric value without forced zero-padding (D11).
    If maxVal = Fix(maxVal) Then
        GetPivotFieldNumericMaxValue = CLng(maxVal)
    Else
        GetPivotFieldNumericMaxValue = maxVal
    End If
End Function

Public Function GetPivotFieldNumericMaxAndContains(ByVal pt As PivotTable, ByVal fieldName As String, ByVal ctx As String, ByRef outMaxVal As Double, ByVal testValue As Double) As Boolean
    ' B3: pre-filter field max and exact domain membership over same PivotItems+TryParseNumericConvertible
    ' enumeration. Membership uses exact CDbl equality semantics via normalized numeric key.
    Dim hasNum As Boolean
    Dim domain As Object
    Dim key As String

    CollectPivotFieldNumericDomain pt, fieldName, ctx, hasNum, outMaxVal, domain
    key = CStr(CDbl(testValue))
    GetPivotFieldNumericMaxAndContains = domain.Exists(key)
End Function

Private Function GetPivotFieldOrientation(ByVal pf As PivotField) As Long
    ' Best-effort orientation read; unexpected errors treated as hidden/unplaced.
    On Error Resume Next
    GetPivotFieldOrientation = pf.Orientation
    If Err.Number <> 0 Then
        Err.Clear
        GetPivotFieldOrientation = xlHidden
    End If
    On Error GoTo 0
End Function

Private Function IsActiveFilterOrientation(ByVal orient As Long) As Boolean
    ' P01/D7: active when already placed in Filters/Columns/Rows.
    IsActiveFilterOrientation = (orient = xlPageField Or orient = xlRowField Or orient = xlColumnField)
End Function

Private Function IsInactiveAvailableOrientation(ByVal orient As Long) As Boolean
    ' P01/D7: available-but-inactive when present in available fields but not placed in Filters/Columns/Rows.
    IsInactiveAvailableOrientation = (orient = xlHidden Or orient = 0)
End Function

Private Sub SortPivotFieldAscending(ByVal pf As PivotField)
    ' P01 best-effort: required sorting when moving available-but-inactive field to Filters/page area.
    ' Regular pivots: attempt ascending label sort. Guarded for unsupported/OLAP cases so runtime continues.
    On Error Resume Next
    Err.Clear
    pf.AutoSort xlAscending, pf.Name
    If Err.Number <> 0 Then
        Err.Clear
        pf.AutoSort xlAscending, pf.SourceName
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub EnsureFilterPlacement(ByVal pf As PivotField)
    ' Q2-G6 / P01: keep field placement and apply filter there if active (page/row/column).
    ' If available but not placed, move to Filters/page and apply. Do not move other active placements.
    Dim orient As Long
    orient = GetPivotFieldOrientation(pf)
    If IsActiveFilterOrientation(orient) Then Exit Sub
    If IsInactiveAvailableOrientation(orient) Then
        On Error Resume Next
        Err.Clear
        pf.Orientation = xlPageField
        If Err.Number = 0 Then
            On Error GoTo 0
            SortPivotFieldAscending pf
            Exit Sub
        End If
        Err.Clear
        On Error GoTo 0
    End If
    ' Unknown/unexpected orientation: preserve existing layout; caller validates filterable placement.
End Sub

Public Function FieldAvailableValues(ByVal pf As PivotField) As Collection
    Dim col As New Collection
    Dim n As String
    Dim i As Long
    On Error Resume Next
    For i = 1 To pf.PivotItems.Count
        n = modTreportConfig.Norm(pf.PivotItems(i).Name)
        If Len(n) > 0 Then
            Err.Clear
            col.Add n, n
            Err.Clear
        End If
    Next i
    On Error GoTo 0
    Set FieldAvailableValues = col
End Function

Public Function FieldAvailableNumericValues(ByVal pf As PivotField) As Collection
    Dim col As New Collection
    Dim n As String
    Dim i As Long
    Dim key As String
    On Error Resume Next
    For i = 1 To pf.PivotItems.Count
        n = modTreportConfig.Norm(pf.PivotItems(i).Name)
        If isNumeric(n) Then
            key = CStr(CDbl(n))
            Err.Clear
            col.Add CDbl(n), "k" & key
            Err.Clear
        End If
    Next i
    On Error GoTo 0
    Set FieldAvailableNumericValues = col
End Function

Public Function MinNumeric(ByVal col As Collection) As Double
    Dim i As Long
    MinNumeric = CDbl(col(1))
    For i = 2 To col.Count
        If CDbl(col(i)) < MinNumeric Then MinNumeric = CDbl(col(i))
    Next i
End Function

Public Function MaxNumeric(ByVal col As Collection) As Double
    Dim i As Long
    MaxNumeric = CDbl(col(1))
    For i = 2 To col.Count
        If CDbl(col(i)) > MaxNumeric Then MaxNumeric = CDbl(col(i))
    Next i
End Function

Public Function CollectionHas(ByVal col As Collection, ByVal key As String) As Boolean
    On Error Resume Next
    Dim v As Variant
    v = col.Item(key)
    CollectionHas = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function JoinDesiredSummary(ByVal desired As Collection) As String
    Dim i As Long
    Dim parts() As String
    Dim tok As String
    Dim n As Long
    If desired Is Nothing Then Exit Function
    If desired.Count = 0 Then Exit Function
    ReDim parts(1 To desired.Count)
    For i = 1 To desired.Count
        tok = modTreportConfig.Norm(CStr(desired(i)))
        n = n + 1
        parts(n) = "'" & tok & "'"
    Next i
    If n = 0 Then Exit Function
    If n < desired.Count Then ReDim Preserve parts(1 To n)
    JoinDesiredSummary = Join(parts, ", ")
End Function

Public Sub ApplyFieldSelection(ByVal pt As PivotTable, ByVal fieldName As String, ByVal desired As Collection, ByVal ctx As String)
    ' Uniform selection rule (T3 / Q2-G4,Q2-G5,Q2-G6,Q2-G9):
    ' - field missing/unavailable in PivotTable fields => fail (2403)
    ' - none of specified values present/available/matchable => fail with clear contextual error (2407)
    ' - at least one specified value present/available/matchable => apply matched subset (no error for unmatched)
    ' Field-scoped operations only; no pt.ClearAllFilters.
    Dim pf As PivotField
    Dim avail As Collection
    Dim wanted As Object
    Dim i As Long
    Dim tok As String, itemName As String
    Dim anyWantedVisible As Boolean
    Dim orient As Long
    Dim matchedKey As String
    Dim reqSummary As String

    Set pf = GetPivotFieldRequired(pt, fieldName, ctx)
    EnsureFilterPlacement pf
    orient = GetPivotFieldOrientation(pf)
    If Not IsActiveFilterOrientation(orient) Then
        modTreportConfig.Fail 2408, ctx & ",field='" & fieldName & "'", "Cannot place or apply filter on field '" & fieldName & "' in a filterable area (Filters/Rows/Columns); placement/apply is impossible."
    End If

    Set avail = FieldAvailableValues(pf)
    Set wanted = CreateObject("Scripting.Dictionary")
    wanted.CompareMode = vbBinaryCompare

    For i = 1 To desired.Count
        tok = modTreportConfig.Norm(CStr(desired(i)))
        If Len(tok) = 0 Then GoTo NextTok
        If CollectionHas(avail, tok) Then
            If Not wanted.Exists(tok) Then wanted.Add tok, True
        Else
            matchedKey = FindMatchingAvailable(avail, tok)
            If Len(matchedKey) > 0 Then
                If Not wanted.Exists(matchedKey) Then wanted.Add matchedKey, True
            End If
            ' Uniform rule: unmatched specified tokens do not cause error when >=1 matches exist.
        End If
NextTok:
    Next i

    If wanted.Count = 0 Then
        reqSummary = JoinDesiredSummary(desired)
        modTreportConfig.Fail 2407, ctx & ",field='" & fieldName & "'", "None of the specified filter values are present or matchable for field '" & fieldName & "'. Requested: [" & reqSummary & "]."
    End If

    On Error GoTo EH
    pt.ManualUpdate = True

    ' Field-scoped clear only (never pt.ClearAllFilters). Preserve layout for active row/column/page fields.
    On Error Resume Next
    pf.ClearAllFilters
    Err.Clear
    pf.ClearManualFilter
    Err.Clear
    If orient = xlPageField Then
        ' Page/filter field: multi-select is page-field behavior; keep existing capability for regular pivots.
        pf.EnableMultiplePageItems = True
        Err.Clear
    End If
    On Error GoTo EH

    ' First pass: ensure at least one wanted item is visible before hiding others.
    For i = 1 To pf.PivotItems.Count
        itemName = modTreportConfig.Norm(pf.PivotItems(i).Name)
        If IsWantedItem(wanted, itemName) Then
            On Error GoTo EH
            pf.PivotItems(i).Visible = True
            anyWantedVisible = True
            Exit For
        End If
    Next i
    If Not anyWantedVisible Then
        reqSummary = JoinDesiredSummary(desired)
        modTreportConfig.Fail 2407, ctx & ",field='" & fieldName & "'", "No valid values remain selected for field '" & fieldName & "' after matching. Requested: [" & reqSummary & "]."
    End If

    ' Second pass: apply matched subset in place without changing field orientation.
    For i = 1 To pf.PivotItems.Count
        itemName = modTreportConfig.Norm(pf.PivotItems(i).Name)
        If IsWantedItem(wanted, itemName) Then
            pf.PivotItems(i).Visible = True
        Else
            On Error Resume Next
            pf.PivotItems(i).Visible = False
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo EH
        End If
    Next i

    pt.ManualUpdate = False
    On Error Resume Next
    pt.Update
    DoEvents
    On Error GoTo 0
    Exit Sub
EH:
    On Error Resume Next
    pt.ManualUpdate = False
    On Error GoTo 0
    modTreportConfig.Fail 2408, ctx & ",field='" & fieldName & "'", "Failed to apply filter selection: " & Err.Description
End Sub

Private Function FindMatchingAvailable(ByVal avail As Collection, ByVal tok As String) As String
    Dim i As Long, a As String, t As String
    t = modTreportConfig.Norm(tok)
    For i = 1 To avail.Count
        a = CStr(avail(i))
        If StrComp(a, t, vbBinaryCompare) = 0 Then
            FindMatchingAvailable = a: Exit Function
        End If
        If isNumeric(a) And isNumeric(t) Then
            If CDbl(a) = CDbl(t) Then
                FindMatchingAvailable = a: Exit Function
            End If
        End If
    Next i
    FindMatchingAvailable = vbNullString
End Function

Private Function IsWantedItem(ByVal wanted As Object, ByVal itemName As String) As Boolean
    If wanted.Exists(itemName) Then
        IsWantedItem = True
        Exit Function
    End If
    Dim k As Variant
    If isNumeric(itemName) Then
        For Each k In wanted.Keys
            If isNumeric(CStr(k)) Then
                If CDbl(CStr(k)) = CDbl(itemName) Then
                    IsWantedItem = True
                    Exit Function
                End If
            End If
        Next k
    End If
    IsWantedItem = False
End Function

Public Function IsPivotDataEmpty(ByVal pt As PivotTable) As Boolean
    On Error Resume Next
    Dim dr As Range
    Set dr = pt.DataBodyRange
    If dr Is Nothing Then
        IsPivotDataEmpty = True
    Else
        IsPivotDataEmpty = (Application.WorksheetFunction.CountA(dr) = 0)
    End If
    On Error GoTo 0
End Function



