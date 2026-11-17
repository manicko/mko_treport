Attribute VB_Name = "modTreportExtract"
Option Explicit

Public Function GetPresentedRange(ByVal pt As PivotTable, ByVal ctx As String) As Range
    Dim rg As Range
    On Error Resume Next
    Set rg = pt.TableRange1
    If rg Is Nothing Then Set rg = pt.TableRange2
    On Error GoTo 0
    If rg Is Nothing Then
        modTreportConfig.Fail 2701, ctx, "Unable to determine presented pivot range."
    End If
    Set GetPresentedRange = rg
End Function

Public Function CopyPivotSnapshot(ByVal srcPt As PivotTable, ByVal dstWs As Worksheet, ByVal pasteRow As Long, ByVal pasteCol As Long, ByVal ctx As String, Optional ByVal onlyValues As Boolean = False, Optional ByVal insertRows As Boolean = False, Optional ByVal skipHeader As Boolean = False, Optional ByVal templateInsertMode As Boolean = False) As Range
    ' P05 (GA-P05-1..7): legacy/out-of-scope insert_rows=TRUE (non-template P09 scope) uses row P=paste_row
    ' as sample/template row. Sequence: copy entire row P and insert H-1 copies starting at P
    ' (H=snapshot.Rows.Count) when H>1, then write snapshot at (P,paste_col). No helper rows/deletion.
    ' Final data P..P+H-1.
    ' P07 (D-P07-1..4): skip_header=TRUE slices presented range to remove column-header band
    ' above DataBodyRange; sliced dimensions drive all downstream operations.
    ' P09 (OP-1..OP-4, 2026-11-17): template-only insert_rows=TRUE (templateInsertMode=True) rework.
    ' Scope inScopeP09 = (insertRows And templateInsertMode): sample row P inserted H-2 times when H>=3,
    ' else 0 when H==2 or H==1 (H==1 allowed). Snapshot pasted at (P-1,paste_col), returned range follows
    ' actual placement (expected P-1..P-1+H-1, A1 verified in L3). Require P>=2 in-scope before any
    ' structural mutation/paste (fail-fast). No deletions, no snapshot shrink. Legacy/non-template
    ' insert_rows=TRUE and insert_rows=FALSE behavior unchanged (Q4=B).
    Dim srcRg As Range
    Dim bodyRg As Range
    Dim dstCell As Range
    Dim pr As Long, pc As Long
    Dim h As Long, w As Long
    Dim hdrOffset As Long
    Dim skipCtx As String
    Dim inScopeP09 As Boolean
    Dim insCount As Long
    Dim dstTop As Long
    Dim scopeCtx As String

    pr = pasteRow
    pc = pasteCol
    If pr <= 0 Then pr = 1
    If pc <= 0 Then pc = 1

    On Error GoTo EH
    On Error Resume Next
    srcPt.ManualUpdate = False
    srcPt.Update
    DoEvents
    On Error GoTo EH

    Set srcRg = GetPresentedRange(srcPt, ctx)

    If skipHeader Then
        skipCtx = ctx & ",data_sheet='" & dstWs.Name & "',parameter='skip_header',operation='skip_header_slicing',paste_row=" & CStr(pr) & ",paste_col=" & CStr(pc) & _
            ",presented_rows=" & CStr(srcRg.Rows.Count) & ",presented_cols=" & CStr(srcRg.Columns.Count) & ",presented_range='" & srcRg.Address(False, False) & "'"

        On Error Resume Next
        Set bodyRg = srcPt.DataBodyRange
        On Error GoTo 0
        If bodyRg Is Nothing Then
            modTreportConfig.Fail 2708, skipCtx & ",body_range='Nothing'", _
                "Cannot apply skip_header: PivotTable.DataBodyRange is unavailable for deterministic slicing."
        End If

        hdrOffset = bodyRg.Row - srcRg.Row
        skipCtx = skipCtx & ",body_row=" & CStr(bodyRg.Row) & ",body_rows=" & CStr(bodyRg.Rows.Count) & ",body_range='" & bodyRg.Address(False, False) & "',header_offset=" & CStr(hdrOffset)

        If hdrOffset < 0 Then
            modTreportConfig.Fail 2708, skipCtx, _
                "Cannot apply skip_header: computed header offset is negative (body.Row - presented.Row = " & CStr(hdrOffset) & ")."
        ElseIf hdrOffset = 0 Then
            ' no-op: presented range already starts at data/value body
        ElseIf hdrOffset >= srcRg.Rows.Count Then
            modTreportConfig.Fail 2709, skipCtx, _
                "Cannot apply skip_header: header offset (" & CStr(hdrOffset) & ") is >= presented row count (" & CStr(srcRg.Rows.Count) & "); sliced snapshot would be empty/invalid."
        Else
            Set srcRg = srcRg.Offset(hdrOffset, 0).Resize(srcRg.Rows.Count - hdrOffset, srcRg.Columns.Count)
        End If

        On Error GoTo EH
    End If

    h = srcRg.Rows.Count
    w = srcRg.Columns.Count
    inScopeP09 = (insertRows And templateInsertMode)
    scopeCtx = ",scope='template_insert_rows_p09',template_insert_mode=" & CStr(templateInsertMode) & ",in_scope_p09=" & CStr(inScopeP09)

    If h <= 0 Or w <= 0 Then
        modTreportConfig.Fail 2704, ctx & ",data_sheet='" & dstWs.Name & "',parameter='insert_rows',paste_row=" & CStr(pr) & ",paste_col=" & CStr(pc) & ",rows=" & CStr(h) & ",cols=" & CStr(w) & ",insert_rows=" & CStr(insertRows) & ",only_values=" & CStr(onlyValues) & ",skip_header=" & CStr(skipHeader) & scopeCtx, _
            "Invalid snapshot size for insertion (rows=" & CStr(h) & ",cols=" & CStr(w) & ")."
    End If

    ' OP-3=A (P09, in-scope only): anchor validation before any structural mutation/paste.
    If inScopeP09 And pr < 2 Then
        modTreportConfig.Fail 2705, ctx & ",data_sheet='" & dstWs.Name & "',parameter='insert_rows/paste_row',paste_row=" & CStr(pr) & ",paste_col=" & CStr(pc) & ",rows=" & CStr(h) & ",cols=" & CStr(w) & ",insert_rows=" & CStr(insertRows) & ",only_values=" & CStr(onlyValues) & ",skip_header=" & CStr(skipHeader) & scopeCtx & ",operation='validate_p09_anchor'", _
            "paste_row must be >= 2 for template-scoped insert_rows (P09 anchor mode) so the snapshot can be placed at paste_row-1; configured/effective paste_row=" & CStr(pr) & "."
    End If

    If inScopeP09 Then
        ' OP-2=B: H>=3 => H-2, H==2 or H==1 => 0 (no fail-fast for H==1). H<=0 already failed.
        If h >= 3 Then
            insCount = h - 2
        Else
            insCount = 0
        End If
        If insCount > 0 Then
            On Error GoTo EHRowPrep
            ' Duplicate sample row P (entire row) H-2 times starting at P; original P shifts down.
            dstWs.Rows(pr).Copy
            dstWs.Rows(pr).Resize(insCount).Insert Shift:=xlDown
            Application.CutCopyMode = False
            On Error GoTo EH
        End If
        dstTop = pr - 1
    ElseIf insertRows And h > 1 Then
        ' Legacy out-of-scope (P05): preserve H-1 inserts at P when H>1.
        insCount = h - 1
        On Error GoTo EHRowPrep
        ' Duplicate sample row P (entire row) H-1 times starting at P; original P shifts down.
        dstWs.Rows(pr).Copy
        dstWs.Rows(pr).Resize(insCount).Insert Shift:=xlDown
        Application.CutCopyMode = False
        On Error GoTo EH
        dstTop = pr
    Else
        ' insert_rows=FALSE or legacy insertRows with h<=1: no row preparation.
        insCount = 0
        dstTop = pr
    End If

    Set dstCell = dstWs.Cells(dstTop, pc)
    PasteSnapshotData srcRg, dstWs, dstCell, onlyValues, ctx

    Set CopyPivotSnapshot = dstWs.Range(dstWs.Cells(dstTop, pc), dstWs.Cells(dstTop + h - 1, pc + w - 1))
    Exit Function

EHRowPrep:
    Application.CutCopyMode = False
    modTreportConfig.Fail 2705, ctx & ",data_sheet='" & dstWs.Name & "',parameter='insert_rows',paste_row=" & CStr(pr) & ",paste_col=" & CStr(pc) & ",rows=" & CStr(h) & ",cols=" & CStr(w) & ",insert_rows=" & CStr(insertRows) & ",only_values=" & CStr(onlyValues) & ",skip_header=" & CStr(skipHeader) & ",insert_count=" & CStr(insCount) & scopeCtx & ",operation='copy/insert sample row'", _
        "Failed to prepare/insert rows for snapshot insertion (sample row copy): " & Err.Description
EH:
    Application.CutCopyMode = False
    modTreportConfig.Fail 2702, ctx & ",data_sheet='" & dstWs.Name & "',parameter='only_values/insert_rows',paste_row=" & CStr(pr) & ",paste_col=" & CStr(pc) & ",rows=" & CStr(h) & ",cols=" & CStr(w) & ",insert_rows=" & CStr(insertRows) & ",only_values=" & CStr(onlyValues) & ",skip_header=" & CStr(skipHeader) & ",dst_top=" & CStr(dstTop) & scopeCtx, _
        "Failed to copy/paste pivot snapshot: " & Err.Description
End Function

Private Sub PasteSnapshotData(ByVal srcRg As Range, ByVal dstWs As Worksheet, ByVal dstCell As Range, ByVal onlyValues As Boolean, ByVal ctx As String)
    Dim i As Long
    Dim h As Long
    Dim dstRg As Range

    h = srcRg.Rows.Count
    On Error GoTo EH
    If onlyValues Then
        ' GA-2: strict values only (no source formats/number formats/widths/heights). Preserve target formatting.
        Set dstRg = dstWs.Range(dstCell, dstCell.Offset(h - 1, srcRg.Columns.Count - 1))
        dstRg.Value = srcRg.Value
        Exit Sub
    End If

    srcRg.Copy
    dstCell.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    srcRg.Copy
    dstCell.PasteSpecial Paste:=xlPasteFormats
    srcRg.Copy
    dstCell.PasteSpecial Paste:=xlPasteColumnWidths
    Application.CutCopyMode = False

    For i = 1 To h
        On Error Resume Next
        dstWs.Rows(dstCell.Row + i - 1).RowHeight = srcRg.Rows(i).RowHeight
        On Error GoTo EH
    Next i
    Exit Sub

EH:
    Application.CutCopyMode = False
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' P05 (GA-P05-7): helper-row deletion removed; insert_rows=TRUE now uses sample row P duplicated with no helper rows.
' P09 (2026-11-17, OP-1..OP-4): template-scoped insert_rows rework uses H-2 inserts (H>=3, else 0 for H==2/H==1)
' and pastes snapshot at P-1 with returned range at actual placement (expected P-1..P-1+H-1, A1 L3-verified).
' Legacy/non-template insert_rows=TRUE retains P05 behavior (H-1 at P when H>1, paste/return at P). No deletions.

Public Sub ApplySheetFormatting(ByVal ws As Worksheet, ByVal freezeRow As Variant, ByVal freezeCol As Variant, ByVal zoomVal As Variant, ByVal ctx As String, Optional ByVal preserveTargetSettings As Boolean = False)
    Dim hasZoom As Boolean
    Dim hasFreeze As Boolean
    Dim fr As Long, fc As Long
    Dim z As Double
    Dim selRow As Long, selCol As Long

    hasZoom = (Not IsEmpty(zoomVal) And Not IsNull(zoomVal) And Len(modTreportConfig.Norm(zoomVal)) > 0)
    fr = 0: fc = 0
    If Not IsEmpty(freezeRow) And Not IsNull(freezeRow) And isNumeric(freezeRow) Then fr = CLng(freezeRow)
    If Not IsEmpty(freezeCol) And Not IsNull(freezeCol) And isNumeric(freezeCol) Then fc = CLng(freezeCol)
    hasFreeze = (fr > 0 Or fc > 0)

    On Error Resume Next
    ws.Activate
    If Not preserveTargetSettings Then
        ActiveWindow.FreezePanes = False
        ActiveWindow.DisplayGridlines = False
    Else
        ' GA-4: preserve template/display settings by default; apply explicit report overrides only.
        If hasFreeze Then ActiveWindow.FreezePanes = False
    End If

    If hasZoom And isNumeric(zoomVal) Then
        z = CDbl(zoomVal)
        If z > 0 And z <= 5 Then z = z * 100
        If z >= 10 And z <= 400 Then ActiveWindow.Zoom = CLng(z)
    End If

    If hasFreeze Then
        selRow = IIf(fr > 0, fr + 1, 1)
        selCol = IIf(fc > 0, fc + 1, 1)
        ws.Cells(selRow, selCol).Select
        ActiveWindow.FreezePanes = True
    End If

    ws.Cells(1, 1).Select
    On Error GoTo 0
End Sub

Private Sub GroupingSkipDiagnostic(ByVal ctx As String, ByVal reason As String)
    Debug.Print "treport[OutlineGrouping][skip][" & ctx & "]: " & reason
End Sub

Public Sub ApplyOutlineGrouping(ByVal insertedRng As Range, ByVal groupRowsRaw As Variant, ByVal rowsCollapsed As Boolean, ByVal ctx As String)
    Dim ws As Worksheet
    Dim groupColRel As Long
    Dim absCol As Long
    Dim insCols As Long
    Dim insRows As Long
    Dim r As Long
    Dim startRow As Long
    Dim inRun As Boolean
    Dim groupRange As Range
    Dim isGroupableFill As Boolean
    Dim cellInterior As Interior
    Dim dblVal As Double
    Dim rawText As String
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean

    If insertedRng Is Nothing Then
        GroupingSkipDiagnostic ctx, "inserted range is Nothing (D6/invalid bounds)."
        Exit Sub
    End If

    Set ws = insertedRng.Worksheet
    insCols = insertedRng.Columns.Count
    insRows = insertedRng.Rows.Count

    ' Deterministic group_rows validation (B4-B2, D6): runtime type/parse/empty/blank/0/missing-equivalent => absent-skip.
    On Error GoTo ValidationTypeErr
    If IsObject(groupRowsRaw) Then
        GroupingSkipDiagnostic ctx & ",group_rows=<object>", "missing/invalid group_rows: object value (D6 skip)."
        Exit Sub
    End If
    If IsError(groupRowsRaw) Then
        GroupingSkipDiagnostic ctx & ",group_rows=<error>", "missing/invalid group_rows: error value (D6 skip)."
        Exit Sub
    End If
    If IsMissing(groupRowsRaw) Or IsEmpty(groupRowsRaw) Or IsNull(groupRowsRaw) Then
        GroupingSkipDiagnostic ctx & ",group_rows=<missing>", "group_rows absent/missing-equivalent (D6 skip)."
        Exit Sub
    End If

    rawText = modTreportConfig.Norm(CStr(groupRowsRaw))
    If Len(rawText) = 0 Then
        GroupingSkipDiagnostic ctx & ",group_rows=<blank>", "group_rows blank/empty (D6 skip)."
        Exit Sub
    End If
    If LCase$(rawText) = "0" Then
        GroupingSkipDiagnostic ctx & ",group_rows=0", "group_rows is 0/missing-equivalent (D6 skip)."
        Exit Sub
    End If
    If Not isNumeric(groupRowsRaw) Then
        GroupingSkipDiagnostic ctx & ",group_rows='" & rawText & "'", "invalid group_rows: non-numeric (D6 skip)."
        Exit Sub
    End If

    dblVal = CDbl(groupRowsRaw)
    If dblVal <> Fix(dblVal) Then
        GroupingSkipDiagnostic ctx & ",group_rows='" & rawText & "'", "invalid group_rows: fractional/non-integer (D6 skip)."
        Exit Sub
    End If
    If dblVal < 1# Or dblVal > 2147483647# Then
        GroupingSkipDiagnostic ctx & ",group_rows='" & rawText & "'", "invalid group_rows: <1 or Long overflow/out-of-range (D6 skip)."
        Exit Sub
    End If
    groupColRel = CLng(dblVal)
    On Error GoTo 0

    If groupColRel > insCols Then
        GroupingSkipDiagnostic ctx & ",group_rows=" & CStr(groupColRel) & ",inserted_width=" & CStr(insCols), "group_rows outside copied range width (D6 skip)."
        Exit Sub
    End If

    absCol = insertedRng.Column + groupColRel - 1

    ' Post-validation grouping: COM/application failures here are real grouping errors (not D6 skips).
    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo GroupingComFail

    ' Ensure outline is available for grouping on the target sheet.
    ws.Outline.SummaryRow = xlSummaryBelow

    inRun = False
    startRow = 0
    For r = 1 To insRows
        Set cellInterior = insertedRng.Cells(r, groupColRel).Interior
        ' Groupable (P01 D5 + compatibility): no fill/pattern-none, or effectively white/system background.
        ' Static Interior only (conditional formatting ignored). Covers explicit white and common
        ' system/automatic/indexed representations (e.g. solid fg indexed 65 in OOXML snapshots).
        isGroupableFill = False
        If cellInterior.Pattern = xlNone Then
            isGroupableFill = True
        ElseIf cellInterior.Pattern = xlSolid Then
            On Error Resume Next
            If CLng(cellInterior.Color) = RGB(255, 255, 255) Then
                isGroupableFill = True
            Else
                Select Case CLng(cellInterior.ColorIndex)
                    Case 2, -4142, -4105, 65
                        ' xlColorIndexWhite (2), xlColorIndexNone (-4142), xlAutomatic (-4105),
                        ' system/default white background (65)
                        isGroupableFill = True
                End Select
            End If
            On Error GoTo 0
        End If

        If isGroupableFill Then
            If Not inRun Then
                startRow = r
                inRun = True
            End If
        Else
            If inRun Then
                If (r - startRow) >= 1 Then
                    Set groupRange = ws.Range(ws.Cells(insertedRng.Row + startRow - 1, absCol), _
                                              ws.Cells(insertedRng.Row + r - 2, absCol))
                    groupRange.Rows.Group
                End If
                inRun = False
            End If
        End If
    Next r

    If inRun Then
        If (insRows - startRow + 1) >= 1 Then
            Set groupRange = ws.Range(ws.Cells(insertedRng.Row + startRow - 1, absCol), _
                                      ws.Cells(insertedRng.Row + insRows - 1, absCol))
            groupRange.Rows.Group
        End If
    End If

    If rowsCollapsed Then
        On Error Resume Next
        ws.Outline.ShowLevels RowLevels:=1
        On Error GoTo GroupingComFail
    End If

    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Exit Sub

ValidationTypeErr:
    GroupingSkipDiagnostic ctx & ",group_rows=<unreadable>", "missing/invalid group_rows: cannot read/parse value (D6 skip). Err=" & CStr(Err.Number) & ": " & Err.Description
    Exit Sub

GroupingComFail:
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    modTreportConfig.Fail 2703, ctx & ",group_rows=" & CStr(groupColRel) & ",abs_col=" & CStr(absCol) & ",rows_collapsed=" & CStr(rowsCollapsed), _
        "Outline grouping failed after validation: " & Err.Description
End Sub




























