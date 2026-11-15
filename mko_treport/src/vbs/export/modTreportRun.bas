Attribute VB_Name = "modTreportRun"
Option Explicit

' treport rev.5 run entry points (FR-RUN-1)
' Final runtime: Excel VBA only. Auxiliary dev tools must not be runtime dependencies.

Private Const OPTIONAL_MISSING_TOKEN As String = "ЗУБНАЯ ПАСТА ДЛЯ ДЕТЕЙ"

Public Sub Refresh_data()
    Dim st As modTreportState.AppState
    Dim cfg As modTreportConfig.TreportConfig
    Dim inputFolder As String
    Dim i As Long
    Dim srcPath As String
    Dim wbSrc As Workbook
    Dim ctx As String

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
    Dim eNum As Long, eDesc As String, eSrc As String
    eNum = Err.Number: eDesc = Err.Description: eSrc = Err.Source
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Application.StatusBar = False
    modTreportState.RestoreState st
    On Error GoTo 0
    MsgBox "Refresh_data failed." & vbCrLf & "Source: " & eSrc & vbCrLf & "Error: " & eNum & vbCrLf & eDesc, vbCritical, "treport"
End Sub

Public Sub Create_report()
    Dim st As modTreportState.AppState
    Dim cfg As modTreportConfig.TreportConfig
    Dim inputFolder As String, outputFolder As String
    Dim i As Long, r As Long
    Dim srcPath As String
    Dim wbSrc As Workbook
    Dim outBooks() As Workbook
    Dim outNames() As String
    Dim ctx As String
    Dim fileCtx As String

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
    Dim outUsedSheetNames() As Object
    ReDim outUsedSheetNames(LBound(cfg.Reports) To UBound(cfg.Reports))

    ' precompute unique trimmed output/report names and create workbooks
    Dim dictNames As Object
    Set dictNames = CreateObject("Scripting.Dictionary")
    For r = LBound(cfg.Reports) To UBound(cfg.Reports)
        Dim rn As String
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

    ' save outputs
    For r = LBound(cfg.Reports) To UBound(cfg.Reports)
        ctx = "Create_report/save[name='" & outNames(r) & "']"
        modTreportOutput.SaveOutputWorkbook outBooks(r), outputFolder, outNames(r), ctx
        Set outBooks(r) = Nothing
    Next r

    Application.StatusBar = False
    modTreportState.RestoreState st
    MsgBox "Create_report completed successfully. Output folder: " & outputFolder, vbInformation, "treport"
    Exit Sub
EH:
    Dim eNum2 As Long, eDesc2 As String, eSrc2 As String
    eNum2 = Err.Number: eDesc2 = Err.Description: eSrc2 = Err.Source
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Dim k As Long
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

    Set ws = modTreportPivot.GetDataSheet(wbSrc, cfg.DataFiles(dfIdx).DataSheet, fileCtx & ",path='" & srcPath & "'")
    Set pt = modTreportPivot.GetPrimaryPivot(ws, fileCtx & ",path='" & srcPath & "'")

    pf = modTreportFilters.ParseAndClassify(cfg.DataFiles(dfIdx).Filters, cfg.General.Delimeter, fileCtx)
    rc = modTreportFilters.ResolveMinMaxAndYear(pt, pf, fileCtx)

    ' Apply data-file filters once (D8)
    modTreportFilters.ApplyDataFileFilters pt, pf, cfg.General.Delimeter, rc, fileCtx

    For r = LBound(cfg.Reports) To UBound(cfg.Reports)
        Dim repCtx As String
        repCtx = fileCtx & ",report='" & cfg.Reports(r).Name & "',report_row=" & cfg.Reports(r).RowIndex

        ' per-report filters: reset/reapply articleLevel4Name_clean and targetAudience
        modTreportFilters.ApplyReportFilters pt, cfg.Reports(r), cfg.General.Delimeter, repCtx

        ' verify no-data? try to detect empty presented data body
        If modTreportPivot.IsPivotDataEmpty(pt) Then
            modTreportConfig.Fail 2410, repCtx, "Filter combination yielded no data."
        End If

        sheetName = modTreportOutput.SubstitutePlaceholders(cfg.DataFiles(dfIdx).ReportSheet, rc.MaxYear, rc.MaxMonth, rc.MaxWeek, repCtx)
        sheetName = modTreportOutput.SanitizeSheetName(sheetName, repCtx)
        ' D14: fail fast on true collision within the same output workbook namespace;
        ' identical final names across different output workbooks are allowed (D-01).
        If outUsedSheetNames(r).Exists(LCase$(sheetName)) Then
            modTreportConfig.Fail 2612, repCtx & ",output='" & outNames(r) & "'", "Sheet name collision in output workbook after substitution/sanitization: '" & sheetName & "'."
        End If
        outUsedSheetNames(r).Add LCase$(sheetName), True

        pr = 1: pc = 1
        If IsNumeric(cfg.DataFiles(dfIdx).pasteRow) And Len(modTreportConfig.Norm(cfg.DataFiles(dfIdx).pasteRow)) > 0 Then pr = CLng(cfg.DataFiles(dfIdx).pasteRow)
        If IsNumeric(cfg.DataFiles(dfIdx).pasteCol) And Len(modTreportConfig.Norm(cfg.DataFiles(dfIdx).pasteCol)) > 0 Then pc = CLng(cfg.DataFiles(dfIdx).pasteCol)
        If pr <= 0 Then pr = 1
        If pc <= 0 Then pc = 1

        Set outWs = modTreportOutput.AddReportSheet(outBooks(r), sheetName, repCtx)
        modTreportExtract.CopyPivotSnapshot pt, outWs, pr, pc, repCtx
        modTreportExtract.ApplySheetFormatting outWs, cfg.DataFiles(dfIdx).freezeRow, cfg.DataFiles(dfIdx).freezeCol, cfg.DataFiles(dfIdx).Zoom, repCtx
        ' D13: group_rows ignored (deferred/no-op)
    Next r
End Sub
