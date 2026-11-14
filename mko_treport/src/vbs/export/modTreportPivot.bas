Attribute VB_Name = "modTreportPivot"
Option Explicit

' treport rev.5 - pivot discovery, members, filter application, presented state

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
    Dim pf As PivotField
    Set pf = TryGetPivotField(pt, fieldName)
    If pf Is Nothing Then
        modTreportConfig.Fail 2403, ctx, "Required pivot field '" & fieldName & "' was not found."
    End If
    Set GetPivotFieldRequired = pf
End Function

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
        If IsNumeric(n) Then
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

Public Sub EnsureFilterPlacement(ByVal pf As PivotField)
    ' D7: do not move if already in Rows/Columns/Filters; add to Filters if hidden
    On Error Resume Next
    Dim o As Long
    o = pf.Orientation
    If Err.Number <> 0 Then
        Err.Clear
        pf.Orientation = xlPageField
    Else
        If o = xlHidden Or o = 0 Then pf.Orientation = xlPageField
    End If
    On Error GoTo 0
End Sub

Public Sub ApplyFieldSelection(ByVal pt As PivotTable, ByVal fieldName As String, ByVal desired As Collection, ByVal allowOptionalSkip As Boolean, ByVal ctx As String)
    Dim pf As PivotField
    Dim avail As Collection
    Dim wanted As Object
    Dim i As Long, j As Long
    Dim tok As String, itemName As String
    Dim anyWantedVisible As Boolean

    Set pf = GetPivotFieldRequired(pt, fieldName, ctx)
    EnsureFilterPlacement pf
    Set avail = FieldAvailableValues(pf)
    Set wanted = CreateObject("Scripting.Dictionary")
    wanted.CompareMode = vbBinaryCompare

    For i = 1 To desired.Count
        tok = modTreportConfig.Norm(desired(i))
        If Len(tok) = 0 Then GoTo NextTok
        If CollectionHas(avail, tok) Then
            If Not wanted.Exists(tok) Then wanted.Add tok, True
        Else
            ' try numeric-normalized match against available
            Dim matchedKey As String
            matchedKey = FindMatchingAvailable(avail, tok)
            If Len(matchedKey) > 0 Then
                If Not wanted.Exists(matchedKey) Then wanted.Add matchedKey, True
            ElseIf allowOptionalSkip And modTreportFilters.IsOptionalSkipToken(tok) Then
                ' skip optional token absent (V-008)
            Else
                modTreportConfig.Fail 2406, ctx & ",field='" & fieldName & "'", "Required filter value '" & tok & "' was not found among available pivot items."
            End If
        End If
NextTok:
    Next i

    If wanted.Count = 0 Then
        modTreportConfig.Fail 2407, ctx & ",field='" & fieldName & "'", "No valid values remain selected for filter."
    End If

    On Error GoTo EH
    pt.ManualUpdate = True
    On Error Resume Next
    pf.EnableMultiplePageItems = True
    pf.ClearAllFilters
    pf.ClearManualFilter
    On Error GoTo EH

    ' Ensure one wanted item visible first
    For i = 1 To pf.PivotItems.Count
        itemName = modTreportConfig.Norm(pf.PivotItems(i).Name)
        If wanted.Exists(itemName) Then
            pf.PivotItems(i).Visible = True
            anyWantedVisible = True
            Exit For
        End If
    Next i
    If Not anyWantedVisible Then
        ' try numeric match by iterating
        For i = 1 To pf.PivotItems.Count
            itemName = modTreportConfig.Norm(pf.PivotItems(i).Name)
            If IsWantedItem(wanted, itemName) Then
                pf.PivotItems(i).Visible = True
                anyWantedVisible = True
                Exit For
            End If
        Next i
    End If
    If Not anyWantedVisible Then modTreportConfig.Fail 2407, ctx & ",field='" & fieldName & "'", "No valid values remain selected for filter (after matching)."

    For i = 1 To pf.PivotItems.Count
        itemName = modTreportConfig.Norm(pf.PivotItems(i).Name)
        If IsWantedItem(wanted, itemName) Then
            pf.PivotItems(i).Visible = True
        Else
            On Error Resume Next
            pf.PivotItems(i).Visible = False
            On Error GoTo EH
        End If
    Next i

    pt.ManualUpdate = False
    On Error Resume Next
    pt.Update
    pt.RefreshTable
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
        If IsNumeric(a) And IsNumeric(t) Then
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
    If IsNumeric(itemName) Then
        For Each k In wanted.Keys
            If IsNumeric(CStr(k)) Then
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
        ' if all cells empty/error/zero-count may still be "shown"; treat as empty only when no data cells
        IsPivotDataEmpty = (dr.Cells.Count = 0)
    End If
    On Error GoTo 0
End Function
