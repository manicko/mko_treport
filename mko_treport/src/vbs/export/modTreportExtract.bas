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

Public Function CopyPivotSnapshot(ByVal srcPt As PivotTable, ByVal dstWs As Worksheet, ByVal pasteRow As Long, ByVal pasteCol As Long, ByVal ctx As String) As Range
    Dim srcRg As Range
    Dim dstCell As Range
    Dim pr As Long, pc As Long
    Dim i As Long

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
    Set dstCell = dstWs.Cells(pr, pc)

    srcRg.Copy
    dstCell.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    srcRg.Copy
    dstCell.PasteSpecial Paste:=xlPasteFormats
    srcRg.Copy
    dstCell.PasteSpecial Paste:=xlPasteColumnWidths
    Application.CutCopyMode = False

    For i = 1 To srcRg.Rows.Count
        On Error Resume Next
        dstWs.Rows(pr + i - 1).RowHeight = srcRg.Rows(i).RowHeight
        On Error GoTo EH
    Next i

    Set CopyPivotSnapshot = dstWs.Range(dstCell, dstCell.Offset(srcRg.Rows.Count - 1, srcRg.Columns.Count - 1))
    Exit Function
EH:
    Application.CutCopyMode = False
    modTreportConfig.Fail 2702, ctx, "Failed to copy/paste pivot snapshot: " & Err.Description
End Function

Public Sub ApplySheetFormatting(ByVal ws As Worksheet, ByVal freezeRow As Variant, ByVal freezeCol As Variant, ByVal zoomVal As Variant, ByVal ctx As String)
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.DisplayGridlines = False
    If Not IsEmpty(zoomVal) And Not IsNull(zoomVal) And Len(modTreportConfig.Norm(zoomVal)) > 0 Then
        If isNumeric(zoomVal) Then
            Dim z As Double
            z = CDbl(zoomVal)
            If z > 0 And z <= 5 Then z = z * 100
            If z >= 10 And z <= 400 Then ActiveWindow.Zoom = CLng(z)
        End If
    End If
    Dim fr As Long, fc As Long
    fr = 0: fc = 0
    If Not IsEmpty(freezeRow) And Not IsNull(freezeRow) And isNumeric(freezeRow) Then fr = CLng(freezeRow)
    If Not IsEmpty(freezeCol) And Not IsNull(freezeCol) And isNumeric(freezeCol) Then fc = CLng(freezeCol)
    If fr > 0 Or fc > 0 Then
        Dim selRow As Long, selCol As Long
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
    Dim isNoFill As Boolean
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
        isNoFill = (cellInterior.Pattern = xlNone) Or _
                   (cellInterior.Pattern = xlSolid And CLng(cellInterior.Color) = RGB(255, 255, 255))

        If isNoFill Then
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
                    If rowsCollapsed Then groupRange.Rows.OutlineLevel = 1
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
            If rowsCollapsed Then groupRange.Rows.OutlineLevel = 1
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
