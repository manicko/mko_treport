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

Private Function IsFreezeParamAbsent(ByVal v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Then
        IsFreezeParamAbsent = True
        Exit Function
    End If
    If IsError(v) Then
        IsFreezeParamAbsent = True
        Exit Function
    End If
    If VarType(v) = vbString Then
        IsFreezeParamAbsent = (Len(Trim$(CStr(v))) = 0)
        Exit Function
    End If
    IsFreezeParamAbsent = False
End Function

Private Function ConfiguredFreezeValue(ByVal v As Variant, ByVal absentAsZero As Boolean) As Long
    If IsFreezeParamAbsent(v) Then
        If absentAsZero Then
            ConfiguredFreezeValue = 0
        Else
            ConfiguredFreezeValue = 0
        End If
        Exit Function
    End If

    If isNumeric(v) Then
        ConfiguredFreezeValue = CLng(v)
        If ConfiguredFreezeValue < 0 Then ConfiguredFreezeValue = 0
    Else
        ConfiguredFreezeValue = 0
    End If
End Function

Public Sub ApplySheetFormatting(ByVal ws As Worksheet, ByVal freezeRow As Variant, ByVal freezeCol As Variant, ByVal zoomVal As Variant, ByVal ctx As String, Optional ByVal preserveTargetSettings As Boolean = False)
    Dim wb As Workbook
    Dim win As Window
    Dim fr As Long, fc As Long
    Dim hasFreeze As Boolean
    Dim hasZoom As Boolean
    Dim rowAbsent As Boolean
    Dim colAbsent As Boolean
    Dim applyFreezeConfig As Boolean
    Dim z As Double
    Dim ctxSheet As String
    Dim ctxFreeze As String

    If ws Is Nothing Then
        modTreportConfig.Fail 2706, ctx & ",operation='apply_sheet_formatting'", "Target worksheet is Nothing."
    End If

    Set wb = ws.Parent
    ctxSheet = ctx & ",data_sheet='" & ws.Name & "'"

    rowAbsent = IsFreezeParamAbsent(freezeRow)
    colAbsent = IsFreezeParamAbsent(freezeCol)
    applyFreezeConfig = (Not preserveTargetSettings) Or (Not rowAbsent) Or (Not colAbsent)

    If applyFreezeConfig Then
        fr = ConfiguredFreezeValue(freezeRow, True)
        fc = ConfiguredFreezeValue(freezeCol, True)
    Else
        fr = 0
        fc = 0
    End If
    hasFreeze = (fr > 0 Or fc > 0)
    ctxFreeze = ctxSheet & ",freeze_row_cfg='" & CStr(freezeRow) & "',freeze_col_cfg='" & CStr(freezeCol) & "',freeze_row_absent=" & CStr(rowAbsent) & ",freeze_col_absent=" & CStr(colAbsent) & ",preserve_target=" & CStr(preserveTargetSettings) & ",apply_freeze_config=" & CStr(applyFreezeConfig) & ",split_row=" & CStr(fr) & ",split_col=" & CStr(fc)

    hasZoom = (Not IsEmpty(zoomVal) And Not IsNull(zoomVal) And Len(modTreportConfig.Norm(zoomVal)) > 0)

    On Error GoTo EH

    ' Activate target workbook/worksheet and verify ActiveWindow belongs to the target context.
    wb.Activate
    ws.Activate

    Set win = ActiveWindow
    If win Is Nothing Then
        modTreportConfig.Fail 2706, ctxFreeze & ",operation='resolve_window'", "ActiveWindow is not available after activating target worksheet."
    End If
    If Not win.Parent Is wb Then
        modTreportConfig.Fail 2706, ctxFreeze & ",operation='resolve_window'", "ActiveWindow does not belong to the target workbook."
    End If
    If Not win.ActiveSheet Is ws Then
        modTreportConfig.Fail 2706, ctxFreeze & ",operation='resolve_window'", "ActiveWindow does not display the target worksheet after activation."
    End If

    ' Deterministic pane baseline and application.
    ' For template sheets with absent freeze settings, preserve existing fixation as-is.
    If applyFreezeConfig Then
        win.FreezePanes = False
        win.Split = False
        On Error Resume Next
        win.SplitRow = 0
        win.SplitColumn = 0
        On Error GoTo EH

        If hasFreeze Then
            win.SplitRow = fr
            win.SplitColumn = fc
            win.FreezePanes = True

            If Not win.FreezePanes Then
                modTreportConfig.Fail 2706, ctxFreeze & ",operation='apply_freeze'", "FreezePanes was not enabled after applying configured split."
            End If
            If win.SplitRow <> fr Then
                modTreportConfig.Fail 2706, ctxFreeze & ",operation='verify_freeze',actual_split_row=" & CStr(win.SplitRow), "Actual SplitRow does not match configured freeze_row."
            End If
            If win.SplitColumn <> fc Then
                modTreportConfig.Fail 2706, ctxFreeze & ",operation='verify_freeze',actual_split_col=" & CStr(win.SplitColumn), "Actual SplitColumn does not match configured freeze_col."
            End If
        Else
            ' Explicit/configured no-freeze state: panes must remain disabled.
            If win.FreezePanes Then
                modTreportConfig.Fail 2706, ctxFreeze & ",operation='verify_freeze',actual_freeze_panes=True", "FreezePanes remained enabled after applying configured no-freeze state."
            End If
            If win.SplitRow <> 0 Then
                modTreportConfig.Fail 2706, ctxFreeze & ",operation='verify_freeze',actual_split_row=" & CStr(win.SplitRow), "SplitRow is not zero after applying configured no-freeze state."
            End If
            If win.SplitColumn <> 0 Then
                modTreportConfig.Fail 2706, ctxFreeze & ",operation='verify_freeze',actual_split_col=" & CStr(win.SplitColumn), "SplitColumn is not zero after applying configured no-freeze state."
            End If
        End If
    End If

    ' Display/zoom overrides: preserve template settings unless explicitly configured/overridden.
    If Not preserveTargetSettings Then
        win.DisplayGridlines = False
    End If

    If hasZoom Then
        If isNumeric(zoomVal) Then
            z = CDbl(zoomVal)
            If z > 0 And z <= 5 Then z = z * 100
            If z >= 10 And z <= 400 Then win.Zoom = CLng(z)
        End If
    End If

    Exit Sub

EH:
    modTreportConfig.Fail 2706, ctxFreeze & ",operation='apply_sheet_formatting'", "Failed to apply sheet formatting deterministically: " & Err.Description
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


' =========================
' B4 (P01): Stage 2 combine — validation, append geometry, transfer fidelity, limits (CR-10..CR-17v, D-02, D-05)
' Hook contract (B3): called for combine groups after Stage 1 all-members preparation, before B5 and before save.
' No temp deletion, no final-registry mutation, no side effects implying success on failure.
' =========================

Private Function B4DictHas(ByVal d As Object, ByVal key As String) As Boolean
    On Error Resume Next
    If d Is Nothing Then
        B4DictHas = False
    Else
        B4DictHas = d.Exists(key)
    End If
    On Error GoTo 0
End Function

Private Function B4DictStr(ByVal d As Object, ByVal key As String, ByVal defaultValue As String) As String
    On Error Resume Next
    If B4DictHas(d, key) Then
        B4DictStr = CStr(d(key))
    Else
        B4DictStr = defaultValue
    End If
    If Err.Number <> 0 Then B4DictStr = defaultValue
    On Error GoTo 0
End Function

Private Function B4DictLong(ByVal d As Object, ByVal key As String, ByVal defaultValue As Long) As Long
    On Error Resume Next
    If B4DictHas(d, key) Then
        B4DictLong = CLng(d(key))
    Else
        B4DictLong = defaultValue
    End If
    If Err.Number <> 0 Then B4DictLong = defaultValue
    On Error GoTo 0
End Function

Private Function B4DictBool(ByVal d As Object, ByVal key As String, ByVal defaultValue As Boolean) As Boolean
    On Error Resume Next
    If B4DictHas(d, key) Then
        B4DictBool = CBool(d(key))
    Else
        B4DictBool = defaultValue
    End If
    If Err.Number <> 0 Then B4DictBool = defaultValue
    On Error GoTo 0
End Function

Private Function B4SafeRangeAddress(ByVal rg As Range) As String
    On Error Resume Next
    If rg Is Nothing Then
        B4SafeRangeAddress = "<Nothing>"
    Else
        B4SafeRangeAddress = rg.Worksheet.Name & "!" & rg.Address(False, False)
    End If
    If Err.Number <> 0 Then B4SafeRangeAddress = "<unavailable>"
    On Error GoTo 0
End Function

Private Function B4MemberCommonContext(ByVal mi As Object, ByVal grp As Object, ByVal outputName As String, ByVal memberIndex As Long) As String
    Dim s As String
    Dim plannedName As String
    Dim baseName As String
    Dim isBase As Boolean
    Dim insertedRg As Range

    plannedName = B4DictStr(mi, "PlannedName", vbNullString)
    baseName = B4DictStr(grp, "BaseName", vbNullString)
    isBase = B4DictBool(mi, "IsBase", (memberIndex = 1))
    s = ",output='" & outputName & "'" & _
        ",final_report_sheet='" & baseName & "'" & _
        ",group_key='" & B4DictStr(grp, "GroupKey", vbNullString) & "'" & _
        ",member_index=" & CStr(memberIndex) & _
        ",is_base=" & CStr(isBase) & _
        ",planned_sheet='" & plannedName & "'" & _
        ",worksheet_name='" & plannedName & "'" & _
        ",associated_report_sheet='" & baseName & "'" & _
        ",report_name='" & B4DictStr(mi, "ReportName", vbNullString) & "'" & _
        ",report_index=" & CStr(B4DictLong(mi, "ReportIndex", 0)) & _
        ",data_file_index=" & CStr(B4DictLong(mi, "DataFileIndex", 0)) & _
        ",data_file_row=" & CStr(B4DictLong(mi, "DataFileRowIndex", 0)) & _
        ",emission_order=" & CStr(B4DictLong(mi, "EmissionOrder", 0)) & _
        ",file_ctx='" & B4DictStr(mi, "FileCtx", vbNullString) & "'" & _
        ",rep_ctx='" & B4DictStr(mi, "RepCtx", vbNullString) & "'" & _
        ",prepared_ctx='" & B4DictStr(mi, "PreparedCtx", vbNullString) & "'"
    On Error Resume Next
    Set insertedRg = Nothing
    If B4DictHas(mi, "InsertedRange") Then Set insertedRg = mi("InsertedRange")
    If Not insertedRg Is Nothing Then s = s & ",inserted_range='" & B4SafeRangeAddress(insertedRg) & "'"
    On Error GoTo 0
    B4MemberCommonContext = s
End Function

Private Function B4GetGroupMember(ByVal grp As Object, ByVal oneBasedIndex As Long, ByVal ctx As String) As Object
    Dim coll As Object
    Dim key As String
    If grp Is Nothing Then
        modTreportConfig.Fail 2698, ctx, "Internal B4 state error: group record is Nothing."
    End If
    If Not B4DictHas(grp, "Members") Then
        modTreportConfig.Fail 2698, ctx, "Internal B4 state error: group record has no Members collection."
    End If
    On Error Resume Next
    Set coll = grp("Members")
    On Error GoTo 0
    If coll Is Nothing Then
        modTreportConfig.Fail 2698, ctx, "Internal B4 state error: group Members collection is Nothing."
    End If
    key = CStr(oneBasedIndex)
    If Not coll.Exists(key) Then
        modTreportConfig.Fail 2698, ctx & ",member_index=" & CStr(oneBasedIndex), "Internal B4 state error: missing member at deterministic index."
    End If
    Set B4GetGroupMember = coll(key)
End Function

Private Sub B4ValidatePreconditions(ByVal outBook As Workbook, ByVal grp As Object, ByVal outputName As String, ByVal reportRowsLimitGeneral As Long, ByVal ctx As String, ByRef memberCount As Long)
    Dim j As Long
    Dim mi As Object
    Dim ws As Worksheet
    Dim plannedName As String

    If outBook Is Nothing Then
        modTreportConfig.Fail 2698, ctx & ",output='" & outputName & "'", "Internal B4 state error: output workbook is Nothing."
    End If
    If grp Is Nothing Then
        modTreportConfig.Fail 2698, ctx & ",output='" & outputName & "'", "Internal B4 state error: group record is Nothing."
    End If
    If Not B4DictHas(grp, "Members") Then
        modTreportConfig.Fail 2698, ctx & ",output='" & outputName & "',group_key='" & B4DictStr(grp, "GroupKey", vbNullString) & "'", "Internal B4 state error: group Members collection missing."
    End If
    On Error Resume Next
    memberCount = grp("Members").Count
    If Err.Number <> 0 Then memberCount = 0
    On Error GoTo 0
    If memberCount < 2 Then
        modTreportConfig.Fail 2698, ctx & ",output='" & outputName & "',effective_base='" & B4DictStr(grp, "BaseName", vbNullString) & "',member_count=" & CStr(memberCount), _
            "Internal B4 state error: combine group requires at least 2 prepared members."
    End If
    If reportRowsLimitGeneral <= 0 Then
        modTreportConfig.Fail 2698, ctx & ",output='" & outputName & "',report_rows_limit_general=" & CStr(reportRowsLimitGeneral), _
            "Internal B4 state error: report_rows_limit_general must be > 0."
    End If

    For j = 1 To memberCount
        Set mi = B4GetGroupMember(grp, j, ctx & ",output='" & outputName & "'")
        If mi Is Nothing Then
            modTreportConfig.Fail 2698, ctx & ",output='" & outputName & "',member_index=" & CStr(j), "Internal B4 state error: member record is Nothing."
        End If
        If Not B4DictBool(mi, "Prepared", False) Then
            modTreportConfig.Fail 2698, ctx & B4MemberCommonContext(mi, grp, outputName, j), "Internal B4 state error: member is not marked Prepared."
        End If
        Set ws = Nothing
        On Error Resume Next
        If B4DictHas(mi, "Sheet") Then Set ws = mi("Sheet")
        On Error GoTo 0
        If ws Is Nothing Then
            modTreportConfig.Fail 2698, ctx & B4MemberCommonContext(mi, grp, outputName, j), "Internal B4 state error: prepared member worksheet is Nothing."
        End If
        plannedName = B4DictStr(mi, "PlannedName", vbNullString)
        If Len(plannedName) = 0 Then
            modTreportConfig.Fail 2698, ctx & B4MemberCommonContext(mi, grp, outputName, j), "Internal B4 state error: member PlannedName is empty."
        End If
        If j = 1 Then
            If Not B4DictBool(mi, "IsBase", False) Then
                modTreportConfig.Fail 2698, ctx & B4MemberCommonContext(mi, grp, outputName, j), "Internal B4 state error: first group member must be base."
            End If
        Else
            If Not B4DictBool(mi, "IsTemp", False) Then
                modTreportConfig.Fail 2698, ctx & B4MemberCommonContext(mi, grp, outputName, j), "Internal B4 state error: non-first group member must be temp."
            End If
        End If
        Set ws = Nothing
        Set mi = Nothing
    Next j
End Sub

Private Function B4RangeHasDataOrFormula(ByVal rg As Range) As Boolean
    ' D-05/CR-17v presence gate: valid content exists if any constant or formula is present in the range.
    ' Accepts valid one-row reports with data/formula; rejects formatting-only/no-content ranges.
    Dim found As Range

    B4RangeHasDataOrFormula = False
    If rg Is Nothing Then Exit Function

    Set found = Nothing
    On Error Resume Next
    Err.Clear
    Set found = rg.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    If Not found Is Nothing Then
        B4RangeHasDataOrFormula = True
        Exit Function
    End If

    Set found = Nothing
    On Error Resume Next
    Err.Clear
    Set found = rg.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If Not found Is Nothing Then
        B4RangeHasDataOrFormula = True
    End If
End Function

Private Sub B4ValidateMemberContentAndLimit(ByVal mi As Object, ByVal grp As Object, ByVal outputName As String, ByVal memberIndex As Long, ByVal reportRowsLimitGeneral As Long, ByVal ctx As String, ByRef sourceBlock As Range)
    Dim ws As Worksheet
    Dim common As String
    Dim actualRows As Long

    Set ws = mi("Sheet")
    Set sourceBlock = ws.UsedRange
    common = ctx & B4MemberCommonContext(mi, grp, outputName, memberIndex) & _
        ",used_range='" & B4SafeRangeAddress(sourceBlock) & "'" & _
        ",used_range_row=" & CStr(sourceBlock.Row) & _
        ",used_range_col=" & CStr(sourceBlock.Column) & _
        ",used_range_rows=" & CStr(sourceBlock.Rows.Count) & _
        ",used_range_cols=" & CStr(sourceBlock.Columns.Count)

    ' D-05 (CR-17v): expected subreport must have valid content range (data or formula). Valid one-row not rejected by row count alone.
    If Not B4RangeHasDataOrFormula(sourceBlock) Then
        modTreportConfig.Fail 2691, common, _
            "Expected subreport has no valid content range (D-05/CR-17v): no data or formulas found in prepared worksheet UsedRange. " & _
            "Empty subreports must not be silently skipped; a valid one-row report is not rejected solely because UsedRange.Rows.Count = 1."
    End If

    ' CR-15/D-02: per-subreport row limit applies to every prepared subreport (base and temps), based on worksheet UsedRange.Rows.Count.
    actualRows = sourceBlock.Rows.Count
    If actualRows > reportRowsLimitGeneral Then
        modTreportConfig.Fail 2692, common & _
            ",actual_row_count=" & CStr(actualRows) & _
            ",configured_report_rows_limit_general=" & CStr(reportRowsLimitGeneral), _
            "Prepared subreport exceeds report_rows_limit_general (CR-15/D-02): worksheet '" & B4DictStr(mi, "PlannedName", vbNullString) & _
            "', associated report_sheet '" & B4DictStr(grp, "BaseName", vbNullString) & _
            "', actual_row_count=" & CStr(actualRows) & _
            ", configured_report_rows_limit_general=" & CStr(reportRowsLimitGeneral) & "."
    End If
End Sub

Private Sub B4ValidateProjectedAppendBounds(ByVal baseWs As Worksheet, ByVal srcBlock As Range, ByVal destinationRow As Long, ByVal gapRows As Long, ByVal baseUsedNow As Range, ByVal mi As Object, ByVal grp As Object, ByVal outputName As String, ByVal memberIndex As Long, ByVal ctx As String)
    Dim projLastRow As Long
    Dim projLastCol As Long
    Dim baseLastRow As Long
    Dim nextRow As Long
    Dim common As String
    Const EXCEL_MAX_ROWS As Long = 1048576
    Const EXCEL_MAX_COLS As Long = 16384

    baseLastRow = baseUsedNow.Row + baseUsedNow.Rows.Count - 1
    nextRow = baseLastRow + 1
    projLastRow = destinationRow + srcBlock.Rows.Count - 1
    projLastCol = srcBlock.Column + srcBlock.Columns.Count - 1

    common = ctx & B4MemberCommonContext(mi, grp, outputName, memberIndex) & _
        ",destination_row=" & CStr(destinationRow) & _
        ",gap_rows=" & CStr(gapRows) & _
        ",base_last_row=" & CStr(baseLastRow) & _
        ",base_next_row=" & CStr(nextRow) & _
        ",base_used_range='" & B4SafeRangeAddress(baseUsedNow) & "'" & _
        ",source_used_range='" & B4SafeRangeAddress(srcBlock) & "'" & _
        ",source_row=" & CStr(srcBlock.Row) & _
        ",source_col=" & CStr(srcBlock.Column) & _
        ",source_rows=" & CStr(srcBlock.Rows.Count) & _
        ",source_cols=" & CStr(srcBlock.Columns.Count)

    ' CR-16: validate projected final combined dimensions before each paste; do not paste if exceeded.
    If destinationRow < 1 Or projLastRow > EXCEL_MAX_ROWS Then
        modTreportConfig.Fail 2693, common & _
            ",projected_row_boundary=" & CStr(projLastRow) & _
            ",applicable_excel_limit=" & CStr(EXCEL_MAX_ROWS), _
            "Projected combined worksheet exceeds Excel row limit before paste (CR-16). " & _
            "final_report_sheet='" & B4DictStr(grp, "BaseName", vbNullString) & _
            "', projected_row_boundary=" & CStr(projLastRow) & _
            ", applicable_excel_limit=" & CStr(EXCEL_MAX_ROWS) & ". Paste was not attempted."
    End If
    If projLastCol > EXCEL_MAX_COLS Then
        modTreportConfig.Fail 2693, common & _
            ",projected_column_boundary=" & CStr(projLastCol) & _
            ",applicable_excel_limit=" & CStr(EXCEL_MAX_COLS), _
            "Projected combined worksheet exceeds Excel column limit before paste (CR-16). " & _
            "final_report_sheet='" & B4DictStr(grp, "BaseName", vbNullString) & _
            "', projected_column_boundary=" & CStr(projLastCol) & _
            ", applicable_excel_limit=" & CStr(EXCEL_MAX_COLS) & ". Paste was not attempted."
    End If
End Sub

Private Sub B4AppendSourceBlock(ByVal srcBlock As Range, ByVal baseWs As Worksheet, ByVal destinationRow As Long, ByVal ctx As String)
    ' CR-12..CR-14: values not formulas; formats incl. number formats; row heights; hidden row states; outline levels at new positions.
    ' CR-13: no column widths, no hidden column states, no zoom/freeze/split/view changes, no base layout/display transfer.
    ' Deterministic object-based operations, no Select/Activate. CutCopyMode cleared (safe-state only) on success/failure.
    Dim srcWs As Worksheet
    Dim dstRg As Range
    Dim dstCell As Range
    Dim rowCount As Long
    Dim colCount As Long
    Dim i As Long
    Dim srcRow As Long
    Dim dstRow As Long

    On Error GoTo EH
    If srcBlock Is Nothing Then
        modTreportConfig.Fail 2694, ctx & ",operation='append_source_block'", "Internal B4 transfer error: source block is Nothing."
    End If
    If baseWs Is Nothing Then
        modTreportConfig.Fail 2694, ctx & ",operation='append_source_block'", "Internal B4 transfer error: base worksheet is Nothing."
    End If

    Set srcWs = srcBlock.Worksheet
    rowCount = srcBlock.Rows.Count
    colCount = srcBlock.Columns.Count
    Set dstCell = baseWs.Cells(destinationRow, srcBlock.Column)
    Set dstRg = baseWs.Range(baseWs.Cells(destinationRow, srcBlock.Column), _
                             baseWs.Cells(destinationRow + rowCount - 1, srcBlock.Column + colCount - 1))

    ' Values only (no formulas). Clipboard-free bulk assignment.
    dstRg.Value = srcBlock.Value

    ' Formats including number formats (not column widths). PasteSpecial xlPasteFormats only.
    srcBlock.Copy
    dstCell.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' Row heights, hidden row states, and outline/grouping levels restored at destination row positions.
    For i = 1 To rowCount
        srcRow = srcBlock.Row + i - 1
        dstRow = destinationRow + i - 1
        baseWs.Rows(dstRow).RowHeight = srcWs.Rows(srcRow).RowHeight
        baseWs.Rows(dstRow).Hidden = srcWs.Rows(srcRow).Hidden
        baseWs.Rows(dstRow).OutlineLevel = srcWs.Rows(srcRow).OutlineLevel
    Next i

    Application.CutCopyMode = False
    Exit Sub

EH:
    On Error Resume Next
    Application.CutCopyMode = False
    On Error GoTo 0
    modTreportConfig.Fail 2694, ctx & _
        ",operation='append_source_block'" & _
        ",destination_row=" & CStr(destinationRow) & _
        ",source_range='" & B4SafeRangeAddress(srcBlock) & "'" & _
        ",row_count=" & CStr(rowCount) & _
        ",col_count=" & CStr(colCount) & _
        ",err_number=" & CStr(Err.Number), _
        "Stage 2 combine transfer failed (CR-12..CR-14): " & Err.Description & _
        ". Clipboard/state cleanup attempted (CutCopyMode cleared); no temp deletion or final-registry changes performed by B4."
End Sub

Public Sub CombinePreparedGroup(ByVal outBook As Workbook, ByVal groupRecord As Object, ByVal outputName As String, ByVal reportRowsLimitGeneral As Long, ByVal ctx As String)
    ' B4 implementation: Stage 2 combine for a prepared combine group (B3 orchestration contract).
    ' - Phase 1: validate/capture UsedRange and enforce D-05/CR-15 for all members (base + temps) before any append.
    ' - Phase 2: append temp members 2..N in deterministic order with CR-10/CR-11 geometry, CR-16 pre-paste bounds,
    '   and CR-12..CR-14 transfer profile (values not formulas, formats incl. number formats, row heights,
    '   hidden row states, outline levels at new positions). CR-13 exclusions preserved. No temp deletion (B5).
    Dim memberCount As Long
    Dim j As Long
    Dim mi As Object
    Dim baseMi As Object
    Dim baseWs As Worksheet
    Dim srcWs As Worksheet
    Dim srcBlock As Range
    Dim baseUsedNow As Range
    Dim sourceBlocks() As Range
    Dim gapRows As Long
    Dim lastRow As Long
    Dim nextRow As Long
    Dim destinationRow As Long
    Dim memberCtx As String

    B4ValidatePreconditions outBook, groupRecord, outputName, reportRowsLimitGeneral, ctx, memberCount

    ReDim sourceBlocks(1 To memberCount)

    ' Phase 1: capture source blocks and validate D-05/CR-15 for all members, including base, before any append/mutation.
    For j = 1 To memberCount
        Set mi = B4GetGroupMember(groupRecord, j, ctx)
        memberCtx = ctx & "/precheck"
        B4ValidateMemberContentAndLimit mi, groupRecord, outputName, j, reportRowsLimitGeneral, memberCtx, srcBlock
        Set sourceBlocks(j) = srcBlock
        Set srcBlock = Nothing
        Set mi = Nothing
    Next j

    Set baseMi = B4GetGroupMember(groupRecord, 1, ctx)
    On Error Resume Next
    Set baseWs = baseMi("Sheet")
    On Error GoTo 0
    If baseWs Is Nothing Then
        modTreportConfig.Fail 2698, ctx & B4MemberCommonContext(baseMi, groupRecord, outputName, 1), _
            "Internal B4 state error: base member worksheet is Nothing during combine."
    End If

    ' Phase 2: append temp members 2..N sequentially in original order; first subreport (base) gap ignored.
    For j = 2 To memberCount
        Set mi = B4GetGroupMember(groupRecord, j, ctx)
        Set srcBlock = sourceBlocks(j)
        If srcBlock Is Nothing Then
            modTreportConfig.Fail 2698, ctx & B4MemberCommonContext(mi, groupRecord, outputName, j), _
                "Internal B4 state error: captured source block is Nothing for append."
        End If
        Set srcWs = srcBlock.Worksheet

        gapRows = B4DictLong(mi, "GapRows", 0)
        If gapRows < 0 Then gapRows = 0

        ' CR-10/CR-11: recompute base geometry from current UsedRange before each append.
        Set baseUsedNow = baseWs.UsedRange
        lastRow = baseUsedNow.Row + baseUsedNow.Rows.Count - 1
        nextRow = lastRow + 1
        destinationRow = nextRow + gapRows

        ' CR-16: projected bounds validated before paste; fail-fast without paste attempt on exceed.
        B4ValidateProjectedAppendBounds baseWs, srcBlock, destinationRow, gapRows, baseUsedNow, mi, groupRecord, outputName, j, ctx & "/bounds"

        ' CR-11/CR-12..CR-14: append preserving source starting column; no paste_row/paste_col reapply.
        B4AppendSourceBlock srcBlock, baseWs, destinationRow, ctx & "/append" & B4MemberCommonContext(mi, groupRecord, outputName, j)

        Set baseUsedNow = Nothing
        Set srcBlock = Nothing
        Set srcWs = Nothing
        Set mi = Nothing
    Next j

    ' Success: no temp deletion/final-registry reconciliation in B4 (B5 responsibility after successful transfer).
    Set baseWs = Nothing
    Set baseMi = Nothing
End Sub

