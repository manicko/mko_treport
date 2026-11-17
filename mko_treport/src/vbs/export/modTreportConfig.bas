Attribute VB_Name = "modTreportConfig"
Option Explicit

' treport rev.5 - config reading & validation (B4)

Public Type TreportGeneral
    inputPath As String
    outputPath As String
    Delimeter As String
    rowsCollapsed As Boolean
End Type

Public Type TreportReport
    ArticleLevel4NameClean As String
    TargetAudience As String
    Name As String
    RowIndex As Long
End Type

Public Type TreportDataFile
    DataFile As String
    DataSheet As String
    ReportSheet As String
    Filters As String
    pasteRow As Variant
    pasteCol As Variant
    freezeRow As Variant
    freezeCol As Variant
    Zoom As Variant
    GroupRows As Variant ' deferred/no-op (D13)
    RowIndex As Long
End Type

Public Type TreportPlaceholder
    Name As String
    Description As String
End Type

Public Type TreportConfig
    General As TreportGeneral
    Reports() As TreportReport
    DataFiles() As TreportDataFile
    Placeholders() As TreportPlaceholder
    HasReports As Boolean
    HasDataFiles As Boolean
End Type

Public Function Norm(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        Norm = vbNullString
    Else
        Norm = Trim$(CStr(v))
    End If
    On Error GoTo 0
End Function

Public Sub Fail(ByVal code As Long, ByVal ctx As String, ByVal msg As String)
    Err.Raise vbObjectError + code, "treport", ctx & ": " & msg
End Sub

Private Function GetSettingsWs(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("settings")
    On Error GoTo 0
    If ws Is Nothing Then Fail 2100, "config", "Required sheet 'settings' was not found."
    Set GetSettingsWs = ws
End Function

Public Function GetLO(ByVal ws As Worksheet, ByVal loName As String) As ListObject
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(loName)
    On Error GoTo 0
    If lo Is Nothing Then
        Fail 2101, "config", "Required ListObject '" & loName & "' was not found on sheet '" & ws.Name & "'."
    End If
    Set GetLO = lo
End Function

Private Function ColIndex(ByVal lo As ListObject, ByVal colName As String) As Long
    Dim lc As ListColumn
    On Error Resume Next
    Set lc = lo.ListColumns(colName)
    On Error GoTo 0
    If lc Is Nothing Then
        Fail 2102, "config", "Required column '" & colName & "' was not found in ListObject '" & lo.Name & "'."
    End If
    ColIndex = lc.Index
End Function

Private Function HasColumn(ByVal lo As ListObject, ByVal colName As String) As Boolean
    Dim lc As ListColumn
    On Error Resume Next
    Set lc = lo.ListColumns(colName)
    On Error GoTo 0
    HasColumn = Not lc Is Nothing
End Function

Public Function CellVal(ByVal lo As ListObject, ByVal rowIdx As Long, ByVal colName As String) As Variant
    Dim ci As Long
    ci = ColIndex(lo, colName)
    If lo.DataBodyRange Is Nothing Then
        CellVal = Empty
    Else
        CellVal = lo.DataBodyRange.Cells(rowIdx, ci).value
    End If
End Function

Public Function CellText(ByVal lo As ListObject, ByVal rowIdx As Long, ByVal colName As String) As String
    CellText = Norm(CellVal(lo, rowIdx, colName))
End Function

Public Function RowCountSafe(ByVal lo As ListObject) As Long
    If lo.DataBodyRange Is Nothing Then
        RowCountSafe = 0
    Else
        RowCountSafe = lo.ListRows.Count
    End If
End Function

Public Function ReadConfig(ByVal wb As Workbook) As TreportConfig
    Dim cfg As TreportConfig
    Dim ws As Worksheet
    Dim loG As ListObject, loR As ListObject, loD As ListObject, loP As ListObject
    Dim i As Long, n As Long
    Dim pName As String, pVal As String
    Dim pValRaw As Variant

    Set ws = GetSettingsWs(wb)
    Set loG = GetLO(ws, "general")
    Set loR = GetLO(ws, "reports")
    Set loD = GetLO(ws, "data_files")
    Set loP = GetLO(ws, "placeholders")

    n = RowCountSafe(loG)
    If n = 0 Then Fail 2103, "config/general", "ListObject 'general' has no data rows."
    cfg.General.rowsCollapsed = False
    For i = 1 To n
        pName = Norm(loG.DataBodyRange.Cells(i, ColIndex(loG, "parameter")).value)
        pValRaw = loG.DataBodyRange.Cells(i, ColIndex(loG, "value")).value
        pVal = Norm(pValRaw)
        Select Case LCase$(pName)
            Case "input_path": cfg.General.inputPath = pVal
            Case "output_path": cfg.General.outputPath = pVal
            Case "delimeter": cfg.General.Delimeter = pVal
            Case "rows_collapsed"
                If IsEmpty(pValRaw) Or IsNull(pValRaw) Then
                    cfg.General.rowsCollapsed = False
                ElseIf VarType(pValRaw) = vbBoolean Then
                    cfg.General.rowsCollapsed = CBool(pValRaw)
                ElseIf Len(pVal) = 0 Then
                    cfg.General.rowsCollapsed = False
                Else
                    Select Case LCase$(pVal)
                        Case "false", "0", "no"
                            cfg.General.rowsCollapsed = False
                        Case "true", "1", "yes"
                            cfg.General.rowsCollapsed = True
                        Case Else
                            Fail 2107, "config/general", "parameter 'rows_collapsed' has invalid value '" & CStr(pValRaw) & "' (row=" & i & ", section='general'); expected blank, False/0/No, or True/1/Yes."
                    End Select
                End If
        End Select
    Next i
    If Len(cfg.General.inputPath) = 0 Then Fail 2104, "config/general", "parameter 'input_path' is required and must be non-empty."
    If Len(cfg.General.outputPath) = 0 Then Fail 2105, "config/general", "parameter 'output_path' is required and must be non-empty."
    If Len(cfg.General.Delimeter) = 0 Then Fail 2106, "config/general", "parameter 'delimeter' is required and must be non-empty."

    n = RowCountSafe(loR)
    If n = 0 Then Fail 2109, "config/reports", "ListObject 'reports' has no data rows."
    ReDim cfg.Reports(1 To n)
    cfg.HasReports = True
    For i = 1 To n
        cfg.Reports(i).ArticleLevel4NameClean = CellText(loR, i, "articleLevel4Name_clean")
        cfg.Reports(i).TargetAudience = CellText(loR, i, "targetAudience")
        cfg.Reports(i).Name = CellText(loR, i, "name")
        cfg.Reports(i).RowIndex = i
        If Len(cfg.Reports(i).Name) = 0 Then Fail 2110, "config/reports[row=" & i & "]", "column 'name' is required."
        If Len(cfg.Reports(i).ArticleLevel4NameClean) = 0 Then Fail 2111, "config/reports[row=" & i & ",name=" & cfg.Reports(i).Name & "]", "column 'articleLevel4Name_clean' is required."
        If Len(cfg.Reports(i).TargetAudience) = 0 Then Fail 2112, "config/reports[row=" & i & ",name=" & cfg.Reports(i).Name & "]", "column 'targetAudience' is required."
    Next i

    n = RowCountSafe(loD)
    If n = 0 Then Fail 2113, "config/data_files", "ListObject 'data_files' has no data rows."
    ReDim cfg.DataFiles(1 To n)
    cfg.HasDataFiles = True
    For i = 1 To n
        cfg.DataFiles(i).DataFile = CellText(loD, i, "data_file")
        cfg.DataFiles(i).DataSheet = CellText(loD, i, "data_sheet")
        cfg.DataFiles(i).ReportSheet = CellText(loD, i, "report_sheet")
        cfg.DataFiles(i).Filters = CellText(loD, i, "filters")
        cfg.DataFiles(i).pasteRow = CellVal(loD, i, "paste_row")
        cfg.DataFiles(i).pasteCol = CellVal(loD, i, "paste_col")
        cfg.DataFiles(i).freezeRow = CellVal(loD, i, "freeze_row")
        cfg.DataFiles(i).freezeCol = CellVal(loD, i, "freeze_col")
        cfg.DataFiles(i).Zoom = CellVal(loD, i, "zoom")
        If HasColumn(loD, "group_rows") Then
            cfg.DataFiles(i).GroupRows = CellVal(loD, i, "group_rows")
        End If
        cfg.DataFiles(i).RowIndex = i
        If Len(cfg.DataFiles(i).DataFile) = 0 Then Fail 2114, "config/data_files[row=" & i & "]", "column 'data_file' is required."
        If Len(cfg.DataFiles(i).DataSheet) = 0 Then Fail 2115, "config/data_files[row=" & i & ",data_file=" & cfg.DataFiles(i).DataFile & "]", "column 'data_sheet' is required."
        If Len(cfg.DataFiles(i).ReportSheet) = 0 Then Fail 2116, "config/data_files[row=" & i & ",data_file=" & cfg.DataFiles(i).DataFile & "]", "column 'report_sheet' is required."
        If Len(cfg.DataFiles(i).Filters) = 0 Then Fail 2117, "config/data_files[row=" & i & ",data_file=" & cfg.DataFiles(i).DataFile & "]", "column 'filters' is required."
    Next i

    n = RowCountSafe(loP)
    If n > 0 Then
        ReDim cfg.Placeholders(1 To n)
        For i = 1 To n
            cfg.Placeholders(i).Name = CellText(loP, i, "placeholder")
            cfg.Placeholders(i).Description = CellText(loP, i, "description")
        Next i
    End If

    ReadConfig = cfg
End Function

Public Function EnsureOutputFolder(ByVal outputPath As String) As String
    Dim p As String
    Dim fso As Object
    p = Norm(outputPath)
    If Len(p) = 0 Then Fail 2120, "config/output_path", "Output path is empty."
    Do While Len(p) > 3 And Right$(p, 1) = "\"
        p = Left$(p, Len(p) - 1)
    Loop
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(p) Then
        On Error GoTo EH
        fso.CreateFolder p
        On Error GoTo 0
    End If
    If Right$(p, 1) <> "\" Then p = p & "\"
    EnsureOutputFolder = p
    Exit Function
EH:
    Fail 2121, "config/output_path='" & outputPath & "'", "Unable to create or access output folder."
End Function

Public Function NormalizeInputFolder(ByVal inputPath As String) As String
    Dim p As String
    Dim fso As Object
    p = Norm(inputPath)
    If Len(p) = 0 Then Fail 2122, "config/input_path", "Input path is empty."
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(p) Then
        Fail 2123, "config/input_path='" & p & "'", "Input folder does not exist."
    End If
    If Right$(p, 1) <> "\" Then p = p & "\"
    NormalizeInputFolder = p
End Function


