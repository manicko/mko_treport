Attribute VB_Name = "modTreportRun"
Option Explicit

' treport rev.6 run entry points (FR-RUN-1)
' Final runtime: Excel VBA only. Auxiliary dev tools must not be runtime dependencies.
' B3 (P01): RC-5 scan/build/preflight preserved; RC-1 2699 gate removed with Stage 1 preparation for all members
' and B4/B5 orchestration hooks. outPlanNames = planned/reserved; outUsedSheetNames = final/base-produced only.

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
    Dim outPlanMembers() As Variant
    Dim outPlanGroups() As Variant
    Dim outPlanNames() As Variant
    Dim planCtx As String
    Dim emissionOrder As Long
    Dim outGroupOrder() As Variant
    Dim outMemberLookup() As Variant

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
    ReDim outPlanMembers(LBound(cfg.DataFiles) To UBound(cfg.DataFiles))
    ReDim outPlanGroups(LBound(cfg.Reports) To UBound(cfg.Reports))
    ReDim outPlanNames(LBound(cfg.Reports) To UBound(cfg.Reports))
    ReDim outGroupOrder(LBound(cfg.Reports) To UBound(cfg.Reports))
    ReDim outMemberLookup(LBound(cfg.Reports) To UBound(cfg.Reports))

    ' RC-5: output workbook objects may be available for inspection, but no planned report worksheets
    ' are created before complete per-run/per-output plan set is resolved, length-validated, and
    ' collision-preflighted (full scan -> build/preflight all plans -> runtime scaffolding -> Stage 1 -> Stage 2 hooks -> save).
    Set dictNames = CreateObject("Scripting.Dictionary")
    dictNames.CompareMode = vbBinaryCompare
    For r = LBound(cfg.Reports) To UBound(cfg.Reports)
        rn = modTreportOutput.SanitizeFileName(cfg.Reports(r).Name, "Create_report/reports[row=" & cfg.Reports(r).RowIndex & "]")
        If dictNames.Exists(LCase$(rn)) Then
            modTreportConfig.Fail 2611, "Create_report/reports[row=" & cfg.Reports(r).RowIndex & ",name='" & cfg.Reports(r).Name & "']", "Report file name collision after Trim/sanitization: '" & rn & "'."
        End If
        dictNames.Add LCase$(rn), True
        outNames(r) = rn
        Set outBooks(r) = Application.Workbooks.Add(xlWBATWorksheet)
        ' RC-3 / B3: final/base-produced registry only (2630/CleanupDefaultSheets/EnsureProducedSheetsPresent/save). Temps never registered.
        Set outUsedSheetNames(r) = CreateObject("Scripting.Dictionary")
        outUsedSheetNames(r).CompareMode = vbBinaryCompare
        Set outPlanGroups(r) = CreateObject("Scripting.Dictionary")
        outPlanGroups(r).CompareMode = vbBinaryCompare
        Set outPlanNames(r) = CreateObject("Scripting.Dictionary")
        outPlanNames(r).CompareMode = vbBinaryCompare
        Set outGroupOrder(r) = Nothing
        Set outMemberLookup(r) = Nothing
    Next r

    ' RC-5 full data_files scan (no report worksheet creation): resolve placeholder max values per df row
    ' before df filters (basis equivalent to current behavior), collect member records, then close source.
    emissionOrder = 0
    For i = LBound(cfg.DataFiles) To UBound(cfg.DataFiles)
        fileCtx = "Create_report[data_file='" & cfg.DataFiles(i).DataFile & "',row=" & cfg.DataFiles(i).RowIndex & ",sheet='" & cfg.DataFiles(i).DataSheet & "']"
        Application.StatusBar = "treport: scan " & fileCtx
        srcPath = modTreportFiles.ResolveDataFilePath(inputFolder, cfg.DataFiles(i).DataFile)
        Set wbSrc = modTreportFiles.OpenWorkbookReadOnly(srcPath)
        CollectDataFilePlanMembers cfg, i, wbSrc, srcPath, fileCtx, outNames, outPlanMembers, emissionOrder
        modTreportFiles.CloseNoSave wbSrc
        Set wbSrc = Nothing
    Next i

    ' Build/preflight all output plans (suffix-aware names/length 2637, collisions 2638, grouping).
    For r = LBound(cfg.Reports) To UBound(cfg.Reports)
        planCtx = "Create_report/plan[name='" & outNames(r) & "',report_row=" & cfg.Reports(r).RowIndex & "]"
        BuildOutputPlan r, outPlanMembers, outPlanGroups(r), outPlanNames(r)
        modTreportOutput.PreflightPlannedSheetNameCollisions outBooks(r), outPlanNames(r), planCtx
    Next r

    ' B3 runtime scaffolding: deterministic group order (first-seen emission) and emission/member-aware lookup.
    ' RC-1 2699 combine-group gate removed and replaced by Stage 1 preparation for all members + B4/B5 hooks.
    For r = LBound(cfg.Reports) To UBound(cfg.Reports)
        BuildOutputRuntimeScaffolding r, outPlanMembers, outPlanGroups(r), outGroupOrder(r), outMemberLookup(r), outNames(r)
    Next r

    ' Stage 1 (B3): prepare all planned members (single-member and combine groups) on B2-assigned sheets.
    ' Preserves Stage 1 semantics/order/contexts: reopen source read-only per df, df filters once, per-report filters,
    ' 2410, placement, PrepareTargetSheet, CopyPivotSnapshot, ApplySheetFormatting, ApplyOutlineGrouping.
    ' Registers only base/final members in outUsedSheetNames(r); temps remain runtime-only.
    For i = LBound(cfg.DataFiles) To UBound(cfg.DataFiles)
        fileCtx = "Create_report[data_file='" & cfg.DataFiles(i).DataFile & "',row=" & cfg.DataFiles(i).RowIndex & ",sheet='" & cfg.DataFiles(i).DataSheet & "']"
        Application.StatusBar = "treport: " & fileCtx
        srcPath = modTreportFiles.ResolveDataFilePath(inputFolder, cfg.DataFiles(i).DataFile)
        Set wbSrc = modTreportFiles.OpenWorkbookReadOnly(srcPath)
        ExecuteDataFilePlannedMembers cfg, i, wbSrc, srcPath, fileCtx, outBooks, outNames, outUsedSheetNames, outMemberLookup
        modTreportFiles.CloseNoSave wbSrc
        Set wbSrc = Nothing
    Next i

    ' Stage 2 (B3 orchestration): B4 combine hook per combine group, then B5 finalization hook only after B4 success.
    ' Hooks are fail-fast-before-save stubs until B4/B5 implemented; EH closes outputs without save.
    For r = LBound(cfg.Reports) To UBound(cfg.Reports)
        ExecuteOutputStage2Hooks r, outBooks(r), outNames(r), outPlanGroups(r), outGroupOrder(r), outUsedSheetNames(r), cfg.General.reportRowsLimitGeneral
    Next r

    ' save outputs (B1/CR-1,D3): cleanup unproduced default sheets on successful normal-completion path only,
    ' with produced-sheet guards and pre/post invariants before save. Temps must be absent and unregistered.
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

Private Sub CollectDataFilePlanMembers(ByRef cfg As modTreportConfig.TreportConfig, ByVal dfIdx As Long, ByVal wbSrc As Workbook, ByVal srcPath As String, ByVal fileCtx As String, ByRef outNames() As String, ByRef outPlanMembers() As Variant, ByRef emissionOrder As Long)
    ' RC-5 scan pass: open source read-only (caller), resolve placeholder max values once per df row before
    ' data_files.filters (equivalent basis), and collect per-emission member records without creating any report worksheet.
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim placeholderMaxValues As Object
    Dim r As Long
    Dim repCtx As String
    Dim substituted As String
    Dim effectiveBase As String
    Dim members() As modTreportOutput.B2SheetMemberPlan
    Dim ub As Long
    Dim hasMembers As Boolean

    Set ws = modTreportPivot.GetDataSheet(wbSrc, cfg.DataFiles(dfIdx).DataSheet, fileCtx & ",path='" & srcPath & "'")
    Set pt = modTreportPivot.GetPrimaryPivot(ws, fileCtx & ",path='" & srcPath & "'")
    Set placeholderMaxValues = ResolveReportSheetPlaceholderMaxValues(pt, cfg.DataFiles(dfIdx).ReportSheet, fileCtx)

    ub = -1
    hasMembers = False
    For r = LBound(cfg.Reports) To UBound(cfg.Reports)
        emissionOrder = emissionOrder + 1
        repCtx = fileCtx & ",report='" & cfg.Reports(r).Name & "',report_row=" & cfg.Reports(r).RowIndex

        substituted = modTreportOutput.SubstituteReportSheetPlaceholders(cfg.DataFiles(dfIdx).ReportSheet, placeholderMaxValues, repCtx)
        effectiveBase = modTreportOutput.ResolveEffectiveBaseSheetName(substituted, repCtx & ",output='" & outNames(r) & "',parameter='report_sheet/data_sheet'", cfg.DataFiles(dfIdx).ReportSheet)

        ub = ub + 1
        If Not hasMembers Then
            ReDim members(0 To 0)
            hasMembers = True
        Else
            ReDim Preserve members(0 To ub)
        End If

        members(ub).ReportIndex = r
        members(ub).DataFileIndex = dfIdx
        members(ub).DataFileRowIndex = cfg.DataFiles(dfIdx).RowIndex
        members(ub).EmissionOrder = emissionOrder
        members(ub).GapRows = cfg.DataFiles(dfIdx).gapRows
        members(ub).EffectiveBaseName = effectiveBase
        members(ub).GroupKey = modTreportOutput.MakeSheetGroupKey(effectiveBase)
        members(ub).PlannedName = vbNullString
        members(ub).SuffixText = vbNullString
        members(ub).SuffixLength = 0
        members(ub).OutputName = outNames(r)
        members(ub).ReportName = cfg.Reports(r).Name
        members(ub).FileCtx = fileCtx
        members(ub).RepCtx = repCtx
        members(ub).RawReportSheetTemplate = cfg.DataFiles(dfIdx).ReportSheet
    Next r

    If hasMembers Then
        outPlanMembers(dfIdx) = members
    Else
        outPlanMembers(dfIdx) = Empty
    End If
End Sub

Private Sub AppendReportMembers(ByRef allMembers() As modTreportOutput.B2SheetMemberPlan, ByRef hasAll As Boolean, ByRef dfMembers() As modTreportOutput.B2SheetMemberPlan, ByVal reportIdx As Long)
    Dim i As Long
    Dim ub As Long
    For i = LBound(dfMembers) To UBound(dfMembers)
        If dfMembers(i).ReportIndex = reportIdx Then
            If Not hasAll Then
                ReDim allMembers(0 To 0)
                hasAll = True
                ub = 0
            Else
                ub = UBound(allMembers) + 1
                ReDim Preserve allMembers(0 To ub)
            End If
            allMembers(UBound(allMembers)) = dfMembers(i)
        End If
    Next i
End Sub

Private Sub BuildOutputPlan(ByVal reportIdx As Long, ByRef dfMembersByDf As Variant, ByVal groupsByKey As Object, ByVal plannedNames As Object)
    ' Build per-output plan in original data_files/emission order across all df rows for this report.
    Dim allMembers() As modTreportOutput.B2SheetMemberPlan
    Dim hasAll As Boolean
    Dim i As Long
    Dim members() As modTreportOutput.B2SheetMemberPlan
    Dim gk As Variant
    Dim grp As Object
    Dim coll As Object
    Dim j As Long
    Dim groupMembers() As modTreportOutput.B2SheetMemberPlan
    Dim mi As Object

    hasAll = False
    For i = LBound(dfMembersByDf) To UBound(dfMembersByDf)
        If IsEmpty(dfMembersByDf(i)) Then GoTo ContinueDf
        members = dfMembersByDf(i)
        AppendReportMembers allMembers, hasAll, members, reportIdx
ContinueDf:
    Next i

    If Not hasAll Then Exit Sub

    ' Stage 1 grouping seed in emission order, then assign planned names per group in that order.
    For i = LBound(allMembers) To UBound(allMembers)
        gk = allMembers(i).GroupKey
        If Not groupsByKey.Exists(CStr(gk)) Then
            Set grp = CreateObject("Scripting.Dictionary")
            grp.CompareMode = vbBinaryCompare
            grp.Add "Key", CStr(gk)
            grp.Add "BaseName", allMembers(i).EffectiveBaseName
            Set coll = CreateObject("Scripting.Dictionary")
            coll.CompareMode = vbBinaryCompare
            grp.Add "Members", coll
            groupsByKey.Add CStr(gk), grp
        End If
        Set grp = groupsByKey(CStr(gk))
        Set coll = grp("Members")
        coll.Add CStr(coll.Count + 1), i
    Next i

    ' Assign suffix-aware planned names and validate lengths (2637) per group in original order.
    For Each gk In groupsByKey.Keys
        Set grp = groupsByKey(CStr(gk))
        Set coll = grp("Members")
        ReDim groupMembers(0 To coll.Count - 1)
        For j = 1 To coll.Count
            groupMembers(j - 1) = allMembers(CLng(coll(CStr(j))))
        Next j
        modTreportOutput.AssignPlannedSheetNames groupMembers
        For j = 1 To coll.Count
            allMembers(CLng(coll(CStr(j)))) = groupMembers(j - 1)
        Next j
    Next gk

    ' Build final planned-name map and replace group member payloads with cloned plan info dictionaries.
    groupsByKey.RemoveAll
    modTreportOutput.BuildOutputSheetPlan allMembers, groupsByKey, plannedNames
End Sub

' B3: RC-1 2699 gate removed; combine groups are handled by Stage 1 preparation for all members and B4/B5 hooks.
' HasCombineGroups/DescribeCombineGroups remain available in modTreportOutput for diagnostics/reporting as needed.

Private Sub BuildOutputRuntimeScaffolding(ByVal reportIdx As Long, ByRef dfMembersByDf As Variant, ByVal groupsByKey As Object, ByRef outGroupOrder As Variant, ByRef outMemberLookup As Variant, ByVal outputName As String)
    ' B3: build deterministic group order (first-seen emission order) and emission/member-aware runtime lookup.
    ' Enriches group member dicts (B2 CloneB2MemberInfo payloads) with runtime fields; does not rely on Dictionary enumeration order.
    Dim lookup As Object
    Dim order() As String
    Dim hasOrder As Boolean
    Dim seen As Object
    Dim i As Long, j As Long
    Dim members() As modTreportOutput.B2SheetMemberPlan
    Dim gk As String
    Dim grp As Object
    Dim coll As Object
    Dim mi As Object
    Dim memberKey As String

    Set lookup = CreateObject("Scripting.Dictionary")
    lookup.CompareMode = vbBinaryCompare
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbBinaryCompare
    hasOrder = False

    ' First-seen group order for this output in original df/report emission order.
    For i = LBound(dfMembersByDf) To UBound(dfMembersByDf)
        If IsEmpty(dfMembersByDf(i)) Then GoTo ContinueDfOrder
        members = dfMembersByDf(i)
        For j = LBound(members) To UBound(members)
            If members(j).ReportIndex <> reportIdx Then GoTo ContinueMemberOrder
            gk = members(j).GroupKey
            If Not seen.Exists(gk) Then
                seen.Add gk, True
                AppendStringToVariantArray order, hasOrder, gk
            End If
ContinueMemberOrder:
        Next j
ContinueDfOrder:
    Next i

    If hasOrder Then
        outGroupOrder = order
    Else
        outGroupOrder = Empty
    End If

    ' Enrich runtime member dicts in B2 group member collections and build emission/member-aware lookup.
    For i = LBound(dfMembersByDf) To UBound(dfMembersByDf)
        If IsEmpty(dfMembersByDf(i)) Then GoTo ContinueDfLookup
        members = dfMembersByDf(i)
        For j = LBound(members) To UBound(members)
            If members(j).ReportIndex <> reportIdx Then GoTo ContinueMemberLookup
            gk = members(j).GroupKey
            If Not groupsByKey.Exists(gk) Then
                modTreportConfig.Fail 2698, "Create_report/runtime_scaffold[output='" & outputName & "',report_index=" & CStr(reportIdx) & ",data_file_index=" & CStr(members(j).DataFileIndex) & ",report_row=" & CStr(members(j).ReportIndex) & "]", _
                    "Internal runtime state error: B2 group not found for planned member."
            End If
            Set grp = groupsByKey(gk)
            Set coll = grp("Members")
            Set mi = FindGroupMemberByEmission(coll, CLng(members(j).EmissionOrder))
            If mi Is Nothing Then
                modTreportConfig.Fail 2698, "Create_report/runtime_scaffold[output='" & outputName & "',emission_order=" & CStr(members(j).EmissionOrder) & "]", _
                    "Internal runtime state error: B2 group member not found for emission."
            End If

            ' Deterministic base/temp identity from B2 member order (1-based in group collection).
            EnsureMemberRuntimeFields mi, grp, outputName
            memberKey = MemberLookupKey(CLng(members(j).DataFileIndex), CLng(members(j).ReportIndex))
            If lookup.Exists(memberKey) Then
                modTreportConfig.Fail 2698, "Create_report/runtime_scaffold[output='" & outputName & "',data_file_index=" & CStr(members(j).DataFileIndex) & ",report_index=" & CStr(members(j).ReportIndex) & "]", _
                    "Internal runtime state error: duplicate emission/member lookup key."
            End If
            lookup.Add memberKey, mi
ContinueMemberLookup:
        Next j
ContinueDfLookup:
    Next i

    Set outMemberLookup = lookup
End Sub

Private Sub AppendStringToVariantArray(ByRef arr() As String, ByRef hasItems As Boolean, ByVal value As String)
    Dim ub As Long
    If Not hasItems Then
        ReDim arr(0 To 0)
        arr(0) = value
        hasItems = True
    Else
        ub = UBound(arr) + 1
        ReDim Preserve arr(0 To ub)
        arr(ub) = value
    End If
End Sub

Private Function MemberLookupKey(ByVal dataFileIndex As Long, ByVal reportIndex As Long) As String
    ' Emission/member-aware key: unique per (df row, report emission) in preserved iteration order.
    MemberLookupKey = CStr(dataFileIndex) & "|" & CStr(reportIndex)
End Function

Private Function FindGroupMemberByEmission(ByVal coll As Object, ByVal emissionOrder As Long) As Object
    Dim k As Variant
    Dim mi As Object
    Set FindGroupMemberByEmission = Nothing
    If coll Is Nothing Then Exit Function
    For Each k In coll.Keys
        Set mi = coll(CStr(k))
        If CLng(mi("EmissionOrder")) = emissionOrder Then
            Set FindGroupMemberByEmission = mi
            Exit Function
        End If
    Next k
End Function

Private Sub EnsureMemberRuntimeFields(ByVal mi As Object, ByVal grp As Object, ByVal outputName As String)
    ' B3 runtime fields on B2 member dicts (deterministic; first member in ordered group collection is base/final).
    Dim memberIndexInGroup As Long
    Dim isBase As Boolean
    Dim coll As Object
    Dim j As Long
    Dim probe As Object

    memberIndexInGroup = 0
    Set coll = grp("Members")
    For j = 1 To coll.Count
        Set probe = coll(CStr(j))
        If CLng(probe("EmissionOrder")) = CLng(mi("EmissionOrder")) Then
            memberIndexInGroup = j
            Exit For
        End If
    Next j
    If memberIndexInGroup <= 0 Then
        modTreportConfig.Fail 2698, "Create_report/runtime_scaffold[output='" & outputName & "',emission_order=" & CStr(mi("EmissionOrder")) & "]", _
            "Internal runtime state error: member index in group not found."
    End If

    isBase = (memberIndexInGroup = 1)

    If Not mi.Exists("MemberIndexInGroup") Then mi.Add "MemberIndexInGroup", memberIndexInGroup Else mi("MemberIndexInGroup") = memberIndexInGroup
    If Not mi.Exists("IsBase") Then mi.Add "IsBase", isBase Else mi("IsBase") = isBase
    If Not mi.Exists("IsTemp") Then mi.Add "IsTemp", (Not isBase) Else mi("IsTemp") = (Not isBase)
    If Not mi.Exists("Sheet") Then mi.Add "Sheet", Nothing
    If Not mi.Exists("InsertedRange") Then mi.Add "InsertedRange", Nothing
    If Not mi.Exists("Prepared") Then mi.Add "Prepared", False Else mi("Prepared") = False
    If Not mi.Exists("CombineStatus") Then mi.Add "CombineStatus", "Pending" Else mi("CombineStatus") = "Pending"
    If Not mi.Exists("PreparedCtx") Then mi.Add "PreparedCtx", vbNullString Else mi("PreparedCtx") = vbNullString
    If Not mi.Exists("OutputNameRuntime") Then mi.Add "OutputNameRuntime", outputName Else mi("OutputNameRuntime") = outputName

    If Not grp.Exists("IsCombineGroup") Then grp.Add "IsCombineGroup", (coll.Count > 1) Else grp("IsCombineGroup") = (coll.Count > 1)
    If Not grp.Exists("CombineStatus") Then grp.Add "CombineStatus", "Pending"
End Sub

Private Sub ExecuteDataFilePlannedMembers(ByRef cfg As modTreportConfig.TreportConfig, ByVal dfIdx As Long, ByVal wbSrc As Workbook, ByVal srcPath As String, ByVal fileCtx As String, ByRef outBooks() As Workbook, ByRef outNames() As String, ByRef outUsedSheetNames() As Object, ByRef outMemberLookup() As Variant)
    ' B3 Stage 1: prepare all planned members (single-member and combine groups) in preserved df/report emission order.
    ' Reopen source read-only per df row (caller), obtain pivot, apply df/report filters in same order,
    ' enforce 2410, compute placement, prepare target with B2-assigned planned name, snapshot/format/grouping,
    ' update runtime member refs, and register only base/final sheets in final-produced registry (RC-3/B3).
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pf As ParsedFilters
    Dim r As Long
    Dim rc As modTreportFilters.ResolvedContext
    Dim repCtx As String
    Dim prepCtx As String
    Dim outWs As Worksheet
    Dim pr As Long, pc As Long
    Dim insertedRng As Range
    Dim plannedName As String
    Dim lookup As Object
    Dim memberKey As String
    Dim mi As Object
    Dim isBase As Boolean

    If IsEmpty(outMemberLookup(LBound(outMemberLookup))) And LBound(outMemberLookup) = UBound(outMemberLookup) Then
        ' no-op guard not reliable for arrays; continue with normal checks per report
    End If

    Set ws = modTreportPivot.GetDataSheet(wbSrc, cfg.DataFiles(dfIdx).DataSheet, fileCtx & ",path='" & srcPath & "'")
    Set pt = modTreportPivot.GetPrimaryPivot(ws, fileCtx & ",path='" & srcPath & "'")

    pf = modTreportFilters.ParseAndClassify(cfg.DataFiles(dfIdx).Filters, cfg.General.Delimeter, fileCtx)
    rc = modTreportFilters.ResolveFilterContext(pt, pf, fileCtx)
    modTreportFilters.ApplyDataFileFilters pt, pf, cfg.General.Delimeter, rc, fileCtx

    For r = LBound(cfg.Reports) To UBound(cfg.Reports)
        repCtx = fileCtx & ",report='" & cfg.Reports(r).Name & "',report_row=" & cfg.Reports(r).RowIndex

        modTreportFilters.ApplyReportFilters pt, cfg.Reports(r), cfg.General.Delimeter, repCtx

        If modTreportPivot.IsPivotDataEmpty(pt) Then
            modTreportConfig.Fail 2410, repCtx, "Filter combination yielded no data."
        End If

        Set lookup = outMemberLookup(r)
        memberKey = MemberLookupKey(dfIdx, r)
        If lookup Is Nothing Or Not lookup.Exists(memberKey) Then
            modTreportConfig.Fail 2698, repCtx & ",output='" & outNames(r) & "',data_file_index=" & CStr(dfIdx) & ",report_index=" & CStr(r), _
                "Internal planning/runtime error: missing B3 runtime member for emission."
        End If
        Set mi = lookup(memberKey)

        plannedName = CStr(mi("PlannedName"))
        If Len(plannedName) = 0 Then
            modTreportConfig.Fail 2698, repCtx & ",output='" & outNames(r) & "'", "Internal planning error: missing B2 planned sheet name for emission."
        End If
        isBase = CBool(mi("IsBase"))
        prepCtx = repCtx & ",output='" & outNames(r) & "',planned_sheet='" & plannedName & "',group_key='" & CStr(mi("GroupKey")) & "',is_base=" & CStr(isBase)

        pr = 1: pc = 1
        If isNumeric(cfg.DataFiles(dfIdx).pasteRow) And Len(modTreportConfig.Norm(cfg.DataFiles(dfIdx).pasteRow)) > 0 Then pr = CLng(cfg.DataFiles(dfIdx).pasteRow)
        If isNumeric(cfg.DataFiles(dfIdx).pasteCol) And Len(modTreportConfig.Norm(cfg.DataFiles(dfIdx).pasteCol)) > 0 Then pc = CLng(cfg.DataFiles(dfIdx).pasteCol)
        If pr <= 0 Then pr = 1
        If pc <= 0 Then pc = 1

        Set outWs = modTreportOutput.PrepareTargetSheet(outBooks(r), plannedName, cfg.DataFiles(dfIdx).template, prepCtx)
        Set insertedRng = modTreportExtract.CopyPivotSnapshot(pt, outWs, pr, pc, prepCtx, cfg.DataFiles(dfIdx).onlyValues, cfg.DataFiles(dfIdx).insertRows, cfg.DataFiles(dfIdx).skipHeader, (Len(modTreportConfig.Norm(cfg.DataFiles(dfIdx).template)) > 0))
        modTreportExtract.ApplySheetFormatting outWs, cfg.DataFiles(dfIdx).freezeRow, cfg.DataFiles(dfIdx).freezeCol, cfg.DataFiles(dfIdx).Zoom, prepCtx, _
            (Len(modTreportConfig.Norm(cfg.DataFiles(dfIdx).template)) > 0)
        modTreportExtract.ApplyOutlineGrouping insertedRng, cfg.DataFiles(dfIdx).GroupRows, cfg.General.rowsCollapsed, prepCtx & ",data_sheet='" & plannedName & "'"

        ' Runtime handoff state for B4/B5 hooks.
        Set mi("Sheet") = outWs
        Set mi("InsertedRange") = insertedRng
        mi("Prepared") = True
        mi("PreparedCtx") = prepCtx
        mi("CombineStatus") = "Prepared"

        ' B3/RC-3: register only base/final sheets in final-produced registry. Temps remain transient runtime-only.
        If isBase Then
            If Not outUsedSheetNames(r).Exists(LCase$(plannedName)) Then
                outUsedSheetNames(r).Add LCase$(plannedName), True
            End If
        End If

        Set insertedRng = Nothing
        Set outWs = Nothing
        Set mi = Nothing
        Set lookup = Nothing
    Next r
End Sub

Private Sub ExecuteOutputStage2Hooks(ByVal reportIdx As Long, ByVal outBook As Workbook, ByVal outputName As String, ByVal groupsByKey As Object, ByRef outGroupOrder As Variant, ByVal finalProducedNames As Object, ByVal reportRowsLimitGeneral As Long)
    ' B3 Stage 2 orchestration: deterministic group order; B4 combine for IsCombineGroup, B5 finalize only after B4 success.
    ' Unimplemented hooks must fail-fast explicitly before cleanup/verify/save; EH closes outputs without save.
    Dim i As Long
    Dim gk As String
    Dim grp As Object
    Dim ctx As String

    If IsEmpty(outGroupOrder) Then Exit Sub
    If groupsByKey Is Nothing Then Exit Sub

    For i = LBound(outGroupOrder) To UBound(outGroupOrder)
        gk = CStr(outGroupOrder(i))
        If Not groupsByKey.Exists(gk) Then
            modTreportConfig.Fail 2698, "Create_report/stage2[output='" & outputName & "',report_index=" & CStr(reportIdx) & ",group_key='" & gk & "']", _
                "Internal runtime state error: group missing during Stage 2 orchestration."
        End If
        Set grp = groupsByKey(gk)
        ctx = "Create_report/stage2[output='" & outputName & "',report_index=" & CStr(reportIdx) & ",group_key='" & gk & "',effective_base='" & CStr(grp("BaseName")) & "']"

        If Not IsCombineGroup(grp) Then
            grp("CombineStatus") = "SingleMember"
            GoTo ContinueGroup
        End If

        ' B4 must complete combine/validations/transfer for group before B5. Stub fails fast before save.
        modTreportExtract.CombinePreparedGroup outBook, grp, outputName, reportRowsLimitGeneral, ctx & "/combine"
        grp("CombineStatus") = "Combined"

        ' B5 must delete temps only after successful transfer and reconcile final/base state. Stub fails fast before save.
        modTreportOutput.FinalizeCombinedGroupTemps outBook, grp, finalProducedNames, ctx & "/finalize"
        grp("CombineStatus") = "Finalized"

ContinueGroup:
        Set grp = Nothing
    Next i
End Sub

Private Function IsCombineGroup(ByVal grp As Object) As Boolean
    If grp Is Nothing Then
        IsCombineGroup = False
        Exit Function
    End If
    If grp.Exists("IsCombineGroup") Then
        IsCombineGroup = CBool(grp("IsCombineGroup"))
        Exit Function
    End If
    If grp.Exists("Members") Then
        IsCombineGroup = (grp("Members").Count > 1)
    Else
        IsCombineGroup = False
    End If
End Function

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
























