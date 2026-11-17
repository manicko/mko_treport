Attribute VB_Name = "modTreportRun"
Option Explicit

' treport rev.6 run entry points (FR-RUN-1)
' Final runtime: Excel VBA only. Auxiliary dev tools must not be runtime dependencies.

Private Const ERR_FILTER_NO_DATA As Long = 2410 + vbObjectError

Public Sub Refresh_data()
    Dim st As modTreportState.AppState
    Dim cfg As modTreportConfig.TreportConfig
    Dim inputFolder As String
    Dim i As Long
    Dim srcPath As String
    Dim wbSrc As Workbook
    Dim ctx As String
    Dim errNumber As Long
    Dim eDesc As String
    Dim eSrc As String

    st = modTreportState.CaptureState()
    On Error GoTo EH
    modTreportState.ApplyFastState st
    Application.StatusBar = "treport: Refresh_data starting..."

    Set wbSrc = Nothing
    cfg = modTreportConfig.ReadConfig(ThisWorkbook)
    inputFolder = modTreportConfig.NormalizeInputFolder(cfg.General.inputPath)
    modTreportConfig.EnsureOutputFolder cfg.General.outputPath

    For i = LBound(cfg.DataFiles) To UBound(cfg.DataFiles)
        ctx = "Refresh_data[data_file='" & cfg.DataFiles(i).DataFile & "',row=" & cfg.DataFiles(i).RowIndex & "]"
        Application.StatusBar = "treport: " & ctx
        srcPath = modTreportFiles.ResolveDataFilePath(inputFolder, cfg.DataFiles(i).DataFile)
        Set wbSrc = modTreportFiles.OpenWorkbookWritable(srcPath)
        modTreportRefresh.RefreshWorkbookData wbSrc, ctx & ",path='" & srcPath & "',sheet='" & cfg.DataFiles(i).DataSheet & "'"
        modTreportFiles.SaveAndClose wbSrc, srcPath
        Set wbSrc = Nothing
    Next i

    Application.StatusBar = False
    modTreportState.RestoreState st
    MsgBox "Refresh_data completed successfully.", vbInformation, "treport"
    Exit Sub
EH:
    errNumber = Err.Number
    eDesc = Err.Description
    eSrc = Err.Source
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Application.StatusBar = False
    modTreportState.RestoreState st
    On Error GoTo 0
    MsgBox "Refresh_data failed." & vbCrLf & "Source: " & eSrc & vbCrLf & "Error: " & errNumber & vbCrLf & eDesc, vbCritical, "treport"
End Sub

Public Sub Create_report()
    Dim st As modTreportState.AppState
    Dim cfg As modTreportConfig.TreportConfig
    Dim inputFolder As String
    Dim outputFolder As String
    Dim i As Long
    Dim r As Long
    Dim srcPath As String
    Dim wbSrc As Workbook
    Dim outBooks() As Workbook
    Dim outNames() As String
    Dim ctx As String
    Dim fileCtx As String
    Dim outUsedSheetNames() As Object
    Dim dictNames As Object
    Dim rn As String
    Dim eNum2 As Long
    Dim eDesc2 As String
    Dim eSrc2 As String
    Dim k As Long

    st = modTreportState.CaptureState()
    On Error GoTo EH
    modTreportState.ApplyFastState st
    Application.StatusBar = "treport: Create_report starting..."

    Set wbSrc = Nothing
    cfg = modTreportConfig.ReadConfig(ThisWorkbook)
    inputFolder = modTreportConfig.NormalizeInputFolder(cfg.General.inputPath)
    outputFolder = modTreportConfig.EnsureOutputFolder(cfg.General.outputPath)

    ReDim outBooks(LBound(cfg.Reports) To UBound(cfg.Reports))
    ReDim outNames(LBound(cfg.Reports) To UBound(cfg.Reports))
    ReDim outUsedSheetNames(LBound(cfg.Reports) To UBound(cfg.Reports))

    ' precompute unique trimmed output/report names and create workbooks
    Set dictNames = CreateObject("Scripting.Dictionary")
    For r = LBound(cfg.Reports) To UBound(cfg.Reports)
        rn = modTreportOutput.SanitizeFileName(cfg.Reports(r).Name, "Create_report/reports[row=" & cfg.Reports(r).RowIndex & "]")
        If dictNames.Exists(LCase$(rn)) Then
            modTreportConfig.Fail 2611, "Create_report/reports[row=" & cfg.Reports(r).RowIndex & ",name='" & cfg.Reports(r).Name & "']", "Report file name collision after Trim/sanitization: '" & rn & "'."
        End If
        dictNames.Add LCase$(rn), True
        outNames(r) = rn
        Set outBooks(r) = Application.Workbooks.Add(xlWBATWorksheet)
        ' Per-output-workbook namespace for final sheet names across all data_files (D3/D14).
        Set outUsedSheetNames(r) = CreateObject("Scripting.Dictionary")
    Next r

    For i = LBound(cfg.DataFiles) To UBound(cfg.DataFiles)
        fileCtx = "Create_report[data_file='" & cfg.DataFiles(i).DataFile & "',row=" & cfg.DataFiles(i).RowIndex & ",sheet='" & cfg.DataFiles(i).DataSheet & "']"
        Application.StatusBar = "treport: " & fileCtx
        srcPath = modTreportFiles.ResolveDataFilePath(inputFolder, cfg.DataFiles(i).DataFile)
        Set wbSrc = modTreportFiles.OpenWorkbookReadOnly(srcPath)
        ProcessDataFile cfg, i, wbSrc, srcPath, fileCtx, outBooks, outNames, outputFolder, outUsedSheetNames
        modTreportFiles.CloseNoSave wbSrc
        Set wbSrc = Nothing
    Next i

    ' save outputs (B1/CR-1,D3): cleanup unproduced default sheets on successful normal-completion path only,
    ' with produced-sheet guards and pre/post invariants before save.
    For r = LBound(cfg.Reports) To UBound(cfg.Reports)
        ctx = "Create_report/save[name='" & outNames(r) & "',report_row=" & cfg.Reports(r).RowIndex & "]"
        If Not outBooks(r) Is Nothing Then
            If outUsedSheetNames(r) Is Nothing Or outUsedSheetNames(r).Count = 0 Then
                modTreportConfig.Fail 2630, ctx, "Output finalization failed: no produced report/content sheets recorded for output workbook; refusing to save invalid workbook."
            End If
            modTreportOutput.CleanupDefaultSheets outBooks(r), outUsedSheetNames(r), ctx & "/cleanup"
            modTreportOutput.EnsureProducedSheetsPresent outBooks(r), outUsedSheetNames(r), ctx & "/verify"
            modTreportOutput.SaveOutputWorkbook outBooks(r), outputFolder, outNames(r), ctx
        End If
        Set outBooks(r) = Nothing
    Next r

    Application.StatusBar = False
    modTreportState.RestoreState st
    MsgBox "Create_report completed successfully. Output folder: " & outputFolder, vbInformation, "treport"
    Exit Sub
EH:
    eNum2 = Err.Number
    eDesc2 = Err.Description
    eSrc2 = Err.Source
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    If (Not Not outBooks) <> 0 Then
        For k = LBound(outBooks) To UBound(outBooks)
            If Not outBooks(k) Is Nothing Then outBooks(k).Close SaveChanges:=False
        Next k
    End If
    Application.StatusBar = False
    modTreportState.RestoreState st
    On Error GoTo 0
    MsgBox "Create_report failed." & vbCrLf & "Source: " & eSrc2 & vbCrLf & "Error: " & eNum2 & vbCrLf & eDesc2, vbCritical, "treport"
End Sub

Private Sub ProcessDataFile(ByRef cfg As modTreportConfig.TreportConfig, ByVal dfIdx As Long, ByVal wbSrc As Workbook, ByVal srcPath As String, ByVal fileCtx As String, ByRef outBooks() As Workbook, ByRef outNames() As String, ByVal outputFolder As String, ByRef outUsedSheetNames() As Object)
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pf As ParsedFilters
    Dim r As Long
    Dim rc As modTreportFilters.ResolvedContext
    Dim sheetName As String
    Dim outWs As Worksheet
    Dim pr As Long, pc As Long
    Dim repCtx As String
    Dim placeholderMaxValues As Object
    Dim insertedRng As Range

    Set ws = modTreportPivot.GetDataSheet(wbSrc, cfg.DataFiles(dfIdx).DataSheet, fileCtx & ",path='" & srcPath & "'")
    Set pt = modTreportPivot.GetPrimaryPivot(ws, fileCtx & ",path='" & srcPath & "'")

    ' B3/D1/D2/D7/D11 (rev.7): resolve report_sheet <max_<field>> and <max-N_<field>> values once per
    ' data_files row from the row pivot/data context (available numeric-convertible field items) BEFORE applying
    ' any data_files.filters or per-report filters, independent of filter clause derivation/representatives.
    Set placeholderMaxValues = ResolveReportSheetPlaceholderMaxValues(pt, cfg.DataFiles(dfIdx).ReportSheet, fileCtx)

    pf = modTreportFilters.ParseAndClassify(cfg.DataFiles(dfIdx).Filters, cfg.General.Delimeter, fileCtx)
    ' T2/D10: resolve global per-field min/max context before applying clauses (independent of other clauses).
    rc = modTreportFilters.ResolveFilterContext(pt, pf, fileCtx)

    ' Apply data-file filters once (D8), deterministic AND in declared clause order.
    modTreportFilters.ApplyDataFileFilters pt, pf, cfg.General.Delimeter, rc, fileCtx

    For r = LBound(cfg.Reports) To UBound(cfg.Reports)
        repCtx = fileCtx & ",report='" & cfg.Reports(r).Name & "',report_row=" & cfg.Reports(r).RowIndex

        ' Per-report filters use the uniform selection rule (no warn-and-skip on none-match).
        modTreportFilters.ApplyReportFilters pt, cfg.Reports(r), cfg.General.Delimeter, repCtx

        ' Fail fast when the filter combination yields no presented data.
        If modTreportPivot.IsPivotDataEmpty(pt) Then
            modTreportConfig.Fail 2410, repCtx, "Filter combination yielded no data."
        End If

        sheetName = modTreportOutput.SubstituteReportSheetPlaceholders(cfg.DataFiles(dfIdx).ReportSheet, placeholderMaxValues, repCtx)
        sheetName = modTreportOutput.SanitizeSheetName(sheetName, repCtx)
        ' D14: fail fast on true collision within the same output workbook namespace;
        ' identical final names across different output workbooks are allowed (D-01).
        If outUsedSheetNames(r).Exists(LCase$(sheetName)) Then
            modTreportConfig.Fail 2612, repCtx & ",output='" & outNames(r) & ",data_sheet='" & sheetName & "',parameter='report_sheet/data_sheet'", _
                "Sheet name collision in output workbook after substitution/sanitization: '" & sheetName & "'."
        End If
        outUsedSheetNames(r).Add LCase$(sheetName), True

        pr = 1: pc = 1
        If isNumeric(cfg.DataFiles(dfIdx).pasteRow) And Len(modTreportConfig.Norm(cfg.DataFiles(dfIdx).pasteRow)) > 0 Then pr = CLng(cfg.DataFiles(dfIdx).pasteRow)
        If isNumeric(cfg.DataFiles(dfIdx).pasteCol) And Len(modTreportConfig.Norm(cfg.DataFiles(dfIdx).pasteCol)) > 0 Then pc = CLng(cfg.DataFiles(dfIdx).pasteCol)
        If pr <= 0 Then pr = 1
        If pc <= 0 Then pc = 1

        Set outWs = modTreportOutput.PrepareTargetSheet(outBooks(r), sheetName, cfg.DataFiles(dfIdx).template, repCtx)
        Set insertedRng = modTreportExtract.CopyPivotSnapshot(pt, outWs, pr, pc, repCtx, cfg.DataFiles(dfIdx).onlyValues, cfg.DataFiles(dfIdx).insertRows, cfg.DataFiles(dfIdx).skipHeader, (Len(modTreportConfig.Norm(cfg.DataFiles(dfIdx).template)) > 0))
        modTreportExtract.ApplySheetFormatting outWs, cfg.DataFiles(dfIdx).freezeRow, cfg.DataFiles(dfIdx).freezeCol, cfg.DataFiles(dfIdx).Zoom, repCtx, _
            (Len(modTreportConfig.Norm(cfg.DataFiles(dfIdx).template)) > 0)
        modTreportExtract.ApplyOutlineGrouping insertedRng, cfg.DataFiles(dfIdx).GroupRows, cfg.General.rowsCollapsed, repCtx & ",output='" & outNames(r) & "',data_sheet='" & sheetName & "'"
        Set insertedRng = Nothing
        Set outWs = Nothing
    Next r
End Sub

Private Function ResolveReportSheetPlaceholderMaxValues(ByVal pt As PivotTable, ByVal reportSheetTemplate As String, ByVal ctx As String) As Object
    ' Computes row-scoped <max_<field>> / <max-N_<field>> values for a data_files row template.
    ' Resolution point: after primary PivotTable is available, before data_files.filters apply.
    ' Source: available numeric-convertible item values of the named PivotTable field in that pivot
    ' context (field-level domain via PivotItems enumeration + TryParseNumericConvertible parity),
    ' not filter representatives and not selected-subset results. D1/D2 enforced fail-fast.
    Dim resolved As Object
    Dim units As Object
    Dim fieldMaxCache As Object
    Dim k As Variant
    Dim unit As Object
    Dim fieldName As String
    Dim nOffset As Long
    Dim maxVal As Double
    Dim computedVal As Double
    Dim hasMember As Boolean
    Dim unitCtx As String
    Dim tokenExample As String
    Dim tokenMap As Object
    Dim valKey As String
    Dim maxText As String
    Dim cacheKey As String
    Dim cacheItem As Object

    Set resolved = CreateObject("Scripting.Dictionary")
    resolved.CompareMode = vbBinaryCompare
    Set fieldMaxCache = CreateObject("Scripting.Dictionary")
    fieldMaxCache.CompareMode = vbBinaryCompare

    Set units = modTreportOutput.CollectMaxPlaceholderUnits(reportSheetTemplate, ctx & ",report_sheet='" & reportSheetTemplate & "'")

    For Each k In units.Keys
        Set unit = units(k)
        fieldName = CStr(unit("Field"))
        nOffset = CLng(unit("N"))
        Set tokenMap = unit("Tokens")
        tokenExample = vbNullString
        If tokenMap.Count > 0 Then
            Dim tk As Variant
            For Each tk In tokenMap.Keys
                tokenExample = CStr(tk)
                Exit For
            Next tk
        End If
        unitCtx = ctx & ",report_sheet='" & reportSheetTemplate & "',template='" & reportSheetTemplate & "',field='" & fieldName & "',N=" & CStr(nOffset) & ",token='" & tokenExample & "'"

        cacheKey = fieldName
        If Not fieldMaxCache.Exists(cacheKey) Then
            computedVal = 0#
            hasMember = modTreportPivot.GetPivotFieldNumericMaxAndContains(pt, fieldName, unitCtx, maxVal, 0#)
            Set cacheItem = CreateObject("Scripting.Dictionary")
            cacheItem.Add "Max", maxVal
            fieldMaxCache.Add cacheKey, cacheItem
        Else
            Set cacheItem = fieldMaxCache(cacheKey)
            maxVal = CDbl(cacheItem("Max"))
        End If

        If nOffset < 0 Then
            modTreportConfig.Fail 2631, unitCtx & ",resolved_max=" & CStr(maxVal), "Invalid <max-N_...> placeholder: N must be >= 0."
        End If
        If CDbl(nOffset) >= maxVal Then
            modTreportConfig.Fail 2632, unitCtx & ",resolved_max=" & CStr(maxVal), "Invalid <max-N_...> placeholder: N must satisfy 0 <= N < max_value (resolved max=" & CStr(maxVal) & ", N=" & CStr(nOffset) & ")."
        End If

        computedVal = maxVal - CDbl(nOffset)
        hasMember = modTreportPivot.GetPivotFieldNumericMaxAndContains(pt, fieldName, unitCtx, maxVal, computedVal)
        If Not hasMember Then
            modTreportConfig.Fail 2633, unitCtx & ",resolved_max=" & CStr(maxVal) & ",computed=" & CStr(computedVal), "Computed <max-N_...> placeholder value is not present in pre-filter field numeric item domain (exact equality)."
        End If

        If computedVal = Fix(computedVal) Then
            maxText = CStr(CLng(computedVal))
        Else
            maxText = CStr(computedVal)
        End If

        valKey = modTreportOutput.MakeMaxPlaceholderUnitKey(fieldName, nOffset)
        If Not resolved.Exists(valKey) Then resolved.Add valKey, maxText
    Next k

    Set ResolveReportSheetPlaceholderMaxValues = resolved
End Function






















