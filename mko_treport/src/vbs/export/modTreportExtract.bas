Attribute VB_Name = "modTreportExtract"
Option Explicit

' treport rev.5 - presented pivot range extraction and static paste (B9,D15)

Public Function GetPresentedRange(ByVal pt As PivotTable, ByVal ctx As String) As Range
    Dim rg As Range
    On Error Resume Next
    ' TableRange1 typically excludes page fields area; TableRange2 may include it.
    Set rg = pt.TableRange1
    If rg Is Nothing Then Set rg = pt.TableRange2
    On Error GoTo 0
    If rg Is Nothing Then
        modTreportConfig.Fail 2701, ctx, "Unable to determine presented pivot range."
    End If
    Set GetPresentedRange = rg
End Function

Public Sub CopyPivotSnapshot(ByVal srcPt As PivotTable, ByVal dstWs As Worksheet, ByVal pasteRow As Long, ByVal pasteCol As Long, ByVal ctx As String)
    Dim srcRg As Range
    Dim dstCell As Range
    Dim pr As Long, pc As Long
    pr = pasteRow
    pc = pasteCol
    If pr <= 0 Then pr = 1
    If pc <= 0 Then pc = 1

    On Error GoTo EH
    srcPt.RefreshTable
    srcPt.Parent.Activate
    Set srcRg = GetPresentedRange(srcPt, ctx)
    srcRg.Copy
    Set dstCell = dstWs.Cells(pr, pc)
    dstCell.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    srcRg.Copy
    dstCell.PasteSpecial Paste:=xlPasteFormats
    srcRg.Copy
    dstCell.PasteSpecial Paste:=xlPasteColumnWidths
    Application.CutCopyMode = False

    ' copy row heights for presented rows
    Dim i As Long
    For i = 1 To srcRg.Rows.Count
        dstWs.Rows(pr + i - 1).RowHeight = srcRg.Rows(i).RowHeight
    Next i
    Exit Sub
EH:
    Application.CutCopyMode = False
    modTreportConfig.Fail 2702, ctx, "Failed to copy/paste pivot snapshot: " & Err.Description
End Sub

Public Sub ApplySheetFormatting(ByVal ws As Worksheet, ByVal freezeRow As Variant, ByVal freezeCol As Variant, ByVal zoomVal As Variant, ByVal ctx As String)
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.DisplayGridlines = False
    If Not IsEmpty(zoomVal) And Not IsNull(zoomVal) And Len(modTreportConfig.Norm(zoomVal)) > 0 Then
        If IsNumeric(zoomVal) Then
            Dim z As Double
            z = CDbl(zoomVal)
            If z > 0 And z <= 5 Then z = z * 100 ' e.g. 0.7 -> 70
            If z >= 10 And z <= 400 Then ActiveWindow.Zoom = CLng(z)
        End If
    End If
    Dim fr As Long, fc As Long
    fr = 0: fc = 0
    If Not IsEmpty(freezeRow) And Not IsNull(freezeRow) And IsNumeric(freezeRow) Then fr = CLng(freezeRow)
    If Not IsEmpty(freezeCol) And Not IsNull(freezeCol) And IsNumeric(freezeCol) Then fc = CLng(freezeCol)
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
