Attribute VB_Name = "xlconcat_macro"
Option Explicit

' ============================================================
' xlconcat
'
' Consolidates data from multiple Excel workbooks into a
' single output workbook.
'
' Entry point:  RunXlconcat  (assigned to run! button)
'
' ThisWorkbook is NEVER closed by this macro.
' ============================================================

' --- Module-level constants ---

Private Const APP_NAME As String = "xlconcat"
Private Const DEFAULT_INTERVAL As Long = 0
Private Const LARGE_USERANGE_THRESHOLD As Long = 10000
Private Const EXCEL_MAX_ROWS As Long = 1048576
Private Const DEBUG_MODE As Boolean = False


' ============================================================
' MAIN ENTRY POINT
' ============================================================

Public Sub RunXlconcat()

    Dim macroWb As Workbook
    Dim outWb As Workbook
    Dim outWs As Worksheet
    Dim wb As Workbook
    Dim ws As Worksheet

    Dim inputPath As String
    Dim outputPath As String

    Dim excludeList As Object
    Dim fileNames As Collection
    Dim f As Variant

    Dim sheetsInterval As Long
    Dim workbookInterval As Long
    Dim ignoreColMismatch As Boolean
    Dim newFileDelimeter As Boolean
    Dim keepColHidden As Boolean

    Dim curPos As Long
    Dim pasteRows As Long
    Dim interval As Long

    Dim srcRange As Range
    Dim tgt As Range

    Dim isFirstSheetInWorkbook As Boolean
    Dim isFirstDataSheet As Boolean
    Dim isFirstSheetOfFile As Boolean

    Dim hasCopiedColumnWidths As Boolean
    Dim hasCopiedZoom As Boolean
    Dim firstVisibleOutputCol As Long

    ' Column-count validation reference (ignore_col_mismatch)
    Dim refColCount As Long
    Dim refFilePath As String
    Dim refSheetName As String

    Dim fileCount As Long
    Dim copiedSheetCount As Long

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim oldDisplayStatusBar As Boolean

    Dim errorNumber As Long
    Dim errorDescription As String
    Dim isError As Boolean


    ' --------------------------------------------------------
    ' Save current Excel settings before anything else.
    ' These are restored in CleanUp even if an error
    ' occurs partway through.
    ' --------------------------------------------------------

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    oldDisplayStatusBar = Application.DisplayStatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False


    Set macroWb = ThisWorkbook


    Debug.Print
    Debug.Print "============================================================"
    Debug.Print "xlconcat START"
    Debug.Print "============================================================"
    Debug.Print "Macro workbook: " & macroWb.FullName


    On Error GoTo FatalError


    ' --------------------------------------------------------
    ' 0. READ SETTINGS
    ' --------------------------------------------------------

    inputPath = GetSettingString(macroWb, "input_path")
    If Len(Trim$(inputPath)) = 0 Then
        Err.Raise vbObjectError + 1000, APP_NAME, _
            "input_path is empty."
    End If

    outputPath = GetSettingString(macroWb, "output_path")
    If Len(Trim$(outputPath)) = 0 Then
        Err.Raise vbObjectError + 1001, APP_NAME, _
            "output_path is empty."
    End If

    sheetsInterval = GetSettingLong(macroWb, "sheets_interval", DEFAULT_INTERVAL)
    workbookInterval = GetSettingLong(macroWb, "workbook_interval", DEFAULT_INTERVAL)

    ignoreColMismatch = GetSettingBool(macroWb, "ignore_col_mismatch")
    newFileDelimeter = GetSettingBool(macroWb, "new_file_delimeter")
    keepColHidden = GetSettingBool(macroWb, "keep_col_hidden")

    If sheetsInterval < 0 Then
        Err.Raise vbObjectError + 1002, APP_NAME, _
            "sheets_interval cannot be negative."
    End If

    If workbookInterval < 0 Then
        Err.Raise vbObjectError + 1003, APP_NAME, _
            "workbook_interval cannot be negative."
    End If


    ' --------------------------------------------------------
    ' Normalize paths
    ' --------------------------------------------------------

    inputPath = NormalizeFolderPath(inputPath)
    outputPath = NormalizeFilePath(outputPath)


    ' --------------------------------------------------------
    ' Output cannot be the macro workbook itself
    ' --------------------------------------------------------

    If PathsEqual(outputPath, macroWb.FullName) Then
        Err.Raise vbObjectError + 1004, APP_NAME, _
            "output_path points to the macro workbook itself:" & _
            vbCrLf & vbCrLf & macroWb.FullName
    End If


    Debug.Print "Settings:"
    Debug.Print "  input_path          = " & inputPath
    Debug.Print "  output_path         = " & outputPath
    Debug.Print "  sheets_interval     = " & sheetsInterval
    Debug.Print "  workbook_interval   = " & workbookInterval
    Debug.Print "  ignore_col_mismatch = " & ignoreColMismatch
    Debug.Print "  new_file_delimeter  = " & newFileDelimeter
    Debug.Print "  keep_col_hidden     = " & keepColHidden


    ' --------------------------------------------------------
    ' Read exclusion list
    ' --------------------------------------------------------

    Set excludeList = ReadExcludeNames()

    Debug.Print "Excluded sheet names: " & excludeList.Count


    ' --------------------------------------------------------
    ' 1. COLLECT INPUT FILES
    ' --------------------------------------------------------

    Debug.Print
    Debug.Print "Collecting input files from: " & inputPath

    Set fileNames = CollectInputFiles(inputPath)

    If fileNames.Count = 0 Then
        MsgBox _
            "In the specified folder:" & vbCrLf & _
            inputPath & vbCrLf & vbCrLf & _
            "no xlsx/xls files found, check settings.", _
            vbCritical, APP_NAME
        GoTo CleanUp
    End If

    Debug.Print "Files found: " & fileNames.Count
    For Each f In fileNames
        Debug.Print "  " & CStr(f)
    Next f


    ' --------------------------------------------------------
    ' 2. CHECK OUTPUT FILE
    ' --------------------------------------------------------

    If FileExists(outputPath) Then
        MsgBox _
            "Output file already exists:" & vbCrLf & vbCrLf & _
            outputPath & vbCrLf & vbCrLf & _
            "Delete it manually and run again.", _
            vbCritical, APP_NAME
        GoTo CleanUp
    End If


    ' --------------------------------------------------------
    ' CREATE OUTPUT WORKBOOK
    ' --------------------------------------------------------

    Set outWb = Workbooks.Add(xlWBATWorksheet)
    Set outWs = outWb.Worksheets(1)
    outWs.name = "plans"

    ' --------------------------------------------------------
    ' Hide gridlines for the output workbook window
    ' --------------------------------------------------------

    outWb.Windows(1).DisplayGridlines = False

    curPos = 1
    fileCount = 0
    copiedSheetCount = 0

    hasCopiedColumnWidths = False
    hasCopiedZoom = False
    isFirstDataSheet = True
    firstVisibleOutputCol = 0


    ' --------------------------------------------------------
    ' 3. PROCESS FILES
    ' --------------------------------------------------------

    For Each f In fileNames
        fileCount = fileCount + 1

        Debug.Print
        Debug.Print "============================================================"
        Debug.Print "FILE #" & fileCount & " : " & CStr(f)
        Debug.Print "============================================================"

        Set wb = Nothing

        Set wb = Workbooks.Open( _
                    fileName:=CStr(f), _
                    UpdateLinks:=0, _
                    ReadOnly:=True, _
                    AddToMru:=False, _
                    IgnoreReadOnlyRecommended:=True)

        Debug.Print "Opened: " & wb.FullName
        Debug.Print "Worksheets: " & wb.Worksheets.Count

        isFirstSheetInWorkbook = True
        isFirstSheetOfFile = True

        For Each ws In wb.Worksheets

            Debug.Print "  Sheet: '" & ws.name & "'  Visible=" & ws.Visible

            ' Skip hidden sheets (xlSheetHidden or xlSheetVeryHidden)
            If ws.Visible = xlSheetVisible Then

                ' Skip excluded names
                If Not excludeList.Exists(CleanName(ws.name)) Then

                    Set srcRange = GetCopyRange(ws)

                    If Not srcRange Is Nothing Then

                        pasteRows = srcRange.Rows.Count

                        If pasteRows > 0 And srcRange.Columns.Count > 0 Then

                            ' --------------------------------
                            ' Column count check (ignore_col_mismatch)
                            ' The first data sheet sets the reference
                            ' column count. Subsequent sheets must match
                            ' unless ignoreColMismatch is True.
                            ' --------------------------------

                            If isFirstDataSheet Then
                                refColCount = srcRange.Columns.Count
                                refFilePath = CStr(f)
                                refSheetName = ws.name
                                Debug.Print "  Reference columns: " & refColCount & _
                                    " (first data sheet: " & refFilePath & "\" & refSheetName & ")"
                            ElseIf Not ignoreColMismatch Then
                                If srcRange.Columns.Count <> refColCount Then
                                    Err.Raise vbObjectError + 1301, APP_NAME, _
                                        "Column count mismatch detected. " & _
                                        "All sheets are pasted to column A of the output, so column counts must be consistent." & _
                                        vbCrLf & vbCrLf & _
                                        "Reference (first copied sheet):" & vbCrLf & _
                                        "  File:   " & refFilePath & vbCrLf & _
                                        "  Sheet:  " & refSheetName & vbCrLf & _
                                        "  Columns: " & refColCount & vbCrLf & vbCrLf & _
                                        "Current sheet:" & vbCrLf & _
                                        "  File:   " & CStr(f) & vbCrLf & _
                                        "  Sheet:  " & ws.name & vbCrLf & _
                                        "  Columns: " & srcRange.Columns.Count & vbCrLf & vbCrLf & _
                                        "Set ignore_col_mismatch = TRUE in the settings table to skip this check."
                                End If
                            End If

                            ' --------------------------------
                            ' Row limit check (spec section 6)
                            ' --------------------------------

                            Dim delimeterRows As Long
                            delimeterRows = 0
                            If newFileDelimeter And isFirstSheetOfFile Then delimeterRows = 1

                            If curPos + interval + delimeterRows + pasteRows - 1 > EXCEL_MAX_ROWS Then
                                Err.Raise vbObjectError + 1300, APP_NAME, _
                                    "Output would exceed Excel's row limit (" & _
                                    EXCEL_MAX_ROWS & " rows). " & _
                                    "Reduce input data or increase sheets_interval."
                            End If

                            ' --------------------------------
                            ' Interval
                            ' --------------------------------

                            If curPos = 1 Then
                                interval = 0
                            ElseIf isFirstSheetInWorkbook Then
                                interval = workbookInterval
                            Else
                                interval = sheetsInterval
                            End If

                            Debug.Print "  Data range: " & srcRange.Address(False, False)
                            Debug.Print "  Paste rows: " & pasteRows
                            Debug.Print "  Interval: " & interval

                            ' --------------------------------
                            ' File delimiter row (optional)
                            ' --------------------------------

                            Dim fileNameOnly As String
                            Dim delimeterRow As Long
                            Dim delimeterWidth As Long
                            Dim outputColStart As Long

                            outputColStart = 1

                            If newFileDelimeter And isFirstSheetOfFile Then
                                delimeterRow = curPos + interval
                                delimeterWidth = srcRange.Columns.Count
                                fileNameOnly = Left$(CStr(f), InStrRev(CStr(f), ".") - 1)
                                fileNameOnly = Mid$(fileNameOnly, InStrRev(fileNameOnly, "\") + 1)

                                ' Determine first visible column (once, for all files)
                                ' Check SOURCE worksheet columns, not output — output
                                ' formatting hasn't been applied yet for the first file.
                                If firstVisibleOutputCol = 0 Then
                                    Dim c As Long
                                    firstVisibleOutputCol = outputColStart
                                    For c = srcRange.Column To srcRange.Column + delimeterWidth - 1
                                        If ws.Columns(c).Hidden = False Then
                                            firstVisibleOutputCol = outputColStart + (c - srcRange.Column)
                                            Exit For
                                        End If
                                    Next c
                                End If

                                With outWs.Range( _
                                    outWs.Cells(delimeterRow, outputColStart), _
                                    outWs.Cells(delimeterRow, outputColStart + delimeterWidth - 1))
                                    outWs.Cells(delimeterRow, firstVisibleOutputCol).value = fileNameOnly
                                    .HorizontalAlignment = xlCenterAcrossSelection
                                    .Interior.Color = RGB(211, 211, 211)
                                    outWs.Cells(delimeterRow, firstVisibleOutputCol).Font.Bold = True
                                    outWs.Cells(delimeterRow, firstVisibleOutputCol).Font.Size = 20
                                End With

                                isFirstSheetOfFile = False
                                Debug.Print "  Target: " & outWs.Cells(curPos + interval + 1, 1).Address
                                Set tgt = outWs.Cells(curPos + interval + 1, 1)
                            Else
                                Debug.Print "  Target: " & outWs.Cells(curPos + interval, 1).Address
                                Set tgt = outWs.Cells(curPos + interval, 1)
                            End If

                            ' --------------------------------
                            ' Stage 1: values + number formats
                            ' --------------------------------

                            Debug.Print "  Copying values + number formats..."

                            srcRange.Copy

                            tgt.PasteSpecial _
                                Paste:=xlPasteValuesAndNumberFormats

                            Application.CutCopyMode = False

                            ' --------------------------------
                            ' Stage 1b: formulas
                            ' Copy formulas that have no external
                            ' workbook reference; external-refer-
                            ' ence cells keep their pasted values.
                            ' --------------------------------

                            Debug.Print "  Processing formulas..."

                            CopyFormulasWithoutExternalLinks _
                                srcRange, tgt

                            ' --------------------------------
                            ' Stage 2: formatting (handles
                            ' merged cells safely)
                            ' --------------------------------

                            Debug.Print "  Copying formats..."

                            srcRange.Copy

                            tgt.PasteSpecial _
                                Paste:=xlPasteFormats

                            Application.CutCopyMode = False

                            ' --------------------------------
                            ' Stage 2: row layout
                            ' --------------------------------

                            ' Row heights (every sheet)
                            CopyRowHeights ws, outWs, _
                                srcRange.Row, srcRange.Rows.Count, tgt.Row

                            ' Hidden rows (applied after heights — spec 2.2)
                            CopyHiddenRows ws, outWs, _
                                srcRange.Row, srcRange.Rows.Count, tgt.Row

                            ' --------------------------------
                            ' Stage 3: first-sheet output layout
                            ' --------------------------------

                            ' Column hidden state control (spec section 8):
                            ' - keepColHidden TRUE: copy the FIRST data
                            '   sheet's column hidden state; subsequent
                            '   sheets must NOT touch it (else they clobber
                            '   the preserved state).
                            ' - keepColHidden FALSE: unhide all columns.
                            ' xlPasteFormats (Stage 2) does not reliably
                            ' copy column hidden state, so we apply it
                            ' explicitly for the first sheet.
                            If isFirstDataSheet Then
                                If keepColHidden Then
                                    CopyColumnHidden ws, outWs, _
                                        srcRange.Column, srcRange.Columns.Count, tgt.Column
                                Else
                                    outWs.Columns.Hidden = False
                                End If
                            ElseIf Not keepColHidden Then
                                outWs.Columns.Hidden = False
                            End If

                            ' Column widths & zoom (first data sheet only)
                            If isFirstDataSheet Then
                                CopyColumnWidths ws, outWs, _
                                    srcRange.Column, srcRange.Columns.Count, tgt.Column
                                hasCopiedColumnWidths = True

                                CopyZoom wb, ws, outWb
                                hasCopiedZoom = True

                                isFirstDataSheet = False
                            End If

                            ' --------------------------------
                            ' Advance output position
                            ' --------------------------------

                            curPos = curPos + interval + delimeterRows + pasteRows
                            isFirstSheetInWorkbook = False
                            copiedSheetCount = copiedSheetCount + 1

                            Debug.Print "  RESULT: copied"

                        End If

                    End If

                End If

            End If

        Next ws

        ' ------------------------------------------------
        ' Close source workbook
        ' ------------------------------------------------

        wb.Close SaveChanges:=False
        Set wb = Nothing

        Debug.Print "Closed: " & CStr(f)

    Next f


    ' --------------------------------------------------------
    ' SUMMARY
    ' --------------------------------------------------------

    Debug.Print
    Debug.Print "============================================================"
    Debug.Print "SUMMARY"
    Debug.Print "============================================================"
    Debug.Print "Files processed: " & fileCount
    Debug.Print "Sheets copied:   " & copiedSheetCount
    Debug.Print "Final output row: " & curPos
    Debug.Print "============================================================"


    ' --------------------------------------------------------
    ' Nothing copied?
    ' --------------------------------------------------------

    If copiedSheetCount = 0 Then
        MsgBox _
            "No sheets were copied." & vbCrLf & vbCrLf & _
            "Files found: " & fileCount & vbCrLf & _
            "Sheets copied: 0" & vbCrLf & vbCrLf & _
            "The output workbook will not be saved." & vbCrLf & _
            "See the VBA Immediate Window (Ctrl+G) for details.", _
            vbExclamation, APP_NAME

        outWb.Close SaveChanges:=False
        Set outWb = Nothing
        GoTo CleanUp
    End If


    ' --------------------------------------------------------
    ' 4. SAVE OUTPUT
    ' --------------------------------------------------------

    Debug.Print "Saving output to: " & outputPath

    outWb.SaveAs _
        fileName:=outputPath, _
        FileFormat:=GetFileFormat(outputPath)

    Debug.Print "Output saved successfully."

    outWb.Close SaveChanges:=True
    Set outWb = Nothing


    ' --------------------------------------------------------
    ' SUCCESS
    ' --------------------------------------------------------

    If DEBUG_MODE Then
        MsgBox _
            "Parameter / Value:" & vbCrLf & _
            "  input_path          = " & inputPath & vbCrLf & _
            "  output_path         = " & outputPath & vbCrLf & _
            "  sheets_interval     = " & sheetsInterval & vbCrLf & _
            "  workbook_interval   = " & workbookInterval & vbCrLf & _
            "  ignore_col_mismatch = " & ignoreColMismatch & vbCrLf & _
            "  new_file_delimeter  = " & newFileDelimeter & vbCrLf & _
            "  keep_col_hidden     = " & keepColHidden, _
            vbInformation, APP_NAME & " DEBUG_MODE"
    End If

    MsgBox _
        "Done." & vbCrLf & vbCrLf & _
        "Files processed: " & fileCount & vbCrLf & _
        "Sheets copied: " & copiedSheetCount & vbCrLf & vbCrLf & _
        "Saved to:" & vbCrLf & _
        outputPath, _
        vbInformation, APP_NAME


    ' --------------------------------------------------------
    ' Normal exit → fall through to CleanUp
    ' --------------------------------------------------------

CleanUp:

    On Error Resume Next

    ' --------------------------------------------------------
    ' Close source workbook (never ThisWorkbook)
    ' --------------------------------------------------------

    If Not wb Is Nothing Then
        If Not wb Is macroWb Then
            wb.Close SaveChanges:=False
        End If
    End If

    ' --------------------------------------------------------
    ' Close output workbook (never ThisWorkbook)
    ' --------------------------------------------------------

    If Not outWb Is Nothing Then
        If Not outWb Is macroWb Then
            outWb.Close SaveChanges:=False
        End If
    End If

    ' --------------------------------------------------------
    ' Clear clipboard
    ' --------------------------------------------------------

    Application.CutCopyMode = False

    ' --------------------------------------------------------
    ' Restore Excel settings
    ' --------------------------------------------------------

    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation
    Application.DisplayStatusBar = oldDisplayStatusBar

    ' --------------------------------------------------------
    ' Show error message (only for unexpected errors)
    ' --------------------------------------------------------

    If isError Then
        MsgBox _
            "xlconcat error." & vbCrLf & vbCrLf & _
            "Error number: " & errorNumber & vbCrLf & _
            "Description: " & errorDescription, _
            vbCritical, APP_NAME
    End If

    Exit Sub


    ' --------------------------------------------------------
    ' FATAL ERROR HANDLER
    ' Captures error details, logs them, and falls
    ' through to CleanUp for resource cleanup.
    ' --------------------------------------------------------

FatalError:

    errorNumber = Err.Number
    errorDescription = Err.Description
    isError = True

    Debug.Print
    Debug.Print "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Debug.Print "FATAL ERROR " & errorNumber
    Debug.Print "Description: " & errorDescription
    Debug.Print "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"

    Resume CleanUp


End Sub


' ============================================================
' SETTINGS — Read from the "settings" table
' Table has two columns: parameter, value
' ============================================================

Private Function GetSettingsTable( _
    ByVal wb As Workbook) As ListObject

    Dim ws As Worksheet
    Dim lo As ListObject

    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets("settings")
    On Error GoTo 0

    If ws Is Nothing Then
        Err.Raise vbObjectError + 1100, APP_NAME, _
            "Settings sheet 'settings' was not found."
    End If

    Set lo = Nothing
    On Error Resume Next
    Set lo = ws.ListObjects("settings")
    On Error GoTo 0

    If lo Is Nothing Then
        Err.Raise vbObjectError + 1101, APP_NAME, _
            "Table 'settings' was not found on the settings sheet."
    End If

    Set GetSettingsTable = lo

End Function


Private Function FindSettingRow( _
    ByVal lo As ListObject, _
    ByVal paramName As String) As ListRow

    Dim rw As ListRow

    For Each rw In lo.ListRows
        If LCase$(Trim$(CStr(rw.Range(1, 1).value))) = LCase$(paramName) Then
            Set FindSettingRow = rw
            Exit Function
        End If
    Next rw

    Set FindSettingRow = Nothing

End Function


Private Function GetSettingString( _
    ByVal wb As Workbook, _
    ByVal paramName As String) As String

    Dim lo As ListObject
    Dim rw As ListRow

    Set lo = GetSettingsTable(wb)
    Set rw = FindSettingRow(lo, paramName)

    If rw Is Nothing Then
        Err.Raise vbObjectError + 1102, APP_NAME, _
            "Parameter '" & paramName & "' was not found in settings table."
    End If

    GetSettingString = Trim$(CStr(rw.Range(1, 2).value))

End Function


Private Function GetSettingLong( _
    ByVal wb As Workbook, _
    ByVal paramName As String, _
    ByVal defaultValue As Long) As Long

    Dim lo As ListObject
    Dim rw As ListRow
    Dim val As Variant

    Set lo = GetSettingsTable(wb)
    Set rw = FindSettingRow(lo, paramName)

    If rw Is Nothing Then
        GetSettingLong = defaultValue
        Exit Function
    End If

    val = rw.Range(1, 2).value

    If Len(Trim$(CStr(val))) = 0 Then
        GetSettingLong = defaultValue
    Else
        GetSettingLong = CLng(val)
    End If

End Function


Private Function GetSettingBool( _
    ByVal wb As Workbook, _
    ByVal paramName As String) As Boolean

    Dim lo As ListObject
    Dim rw As ListRow
    Dim val As Variant

    Set lo = GetSettingsTable(wb)
    Set rw = FindSettingRow(lo, paramName)

    If rw Is Nothing Then
        GetSettingBool = False
        Exit Function
    End If

    val = rw.Range(1, 2).value

    If CBool(LCase$(Trim$(CStr(val))) = "true") Then
        GetSettingBool = True
    Else
        GetSettingBool = False
    End If

End Function


' ============================================================
' EXCLUSION LIST — Read the exclude_names ListObject
' ============================================================

Private Function ReadExcludeNames() As Object

    Dim dict As Object
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim rng As Range
    Dim cell As Range
    Dim key As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("settings")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ReadExcludeNames = dict
        Exit Function
    End If

    Set lo = Nothing
    On Error Resume Next
    Set lo = ws.ListObjects("exclude_names")
    On Error GoTo 0

    If lo Is Nothing Then
        Debug.Print "  Exclude table 'exclude_names' not found."
        Set ReadExcludeNames = dict
        Exit Function
    End If

    Set rng = lo.DataBodyRange
    If rng Is Nothing Then
        Debug.Print "  Exclude table is empty."
        Set ReadExcludeNames = dict
        Exit Function
    End If

    For Each cell In rng.Cells
        key = CleanName(CStr(cell.value))
        If Len(key) > 0 Then
            If Not dict.Exists(key) Then
                dict.Add key, True
            End If
        End If
    Next cell

    Set ReadExcludeNames = dict

End Function


' ============================================================
' SHEET NAME NORMALISATION
' Lowercase, spaces removed (for exclusion comparison)
' ============================================================

Private Function CleanName( _
    ByVal name As String) As String

    Dim s As String

    s = LCase$(Trim$(name))
    s = Replace(s, " ", "")

    CleanName = s

End Function


' ============================================================
' FILE COLLECTION — FileSystemObject
' Searches for .xlsx and .xls files, sorted alphabetically.
' ============================================================

Private Function CollectInputFiles( _
    ByVal inputPath As String) As Collection

    Dim result As New Collection
    Dim fso As Object
    Dim folder As Object
    Dim file As Object

    Dim fileName As String
    Dim extension As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(inputPath) Then
        Err.Raise vbObjectError + 1200, APP_NAME, _
            "Input folder does not exist:" & vbCrLf & inputPath
    End If

    Set folder = fso.GetFolder(inputPath)

    For Each file In folder.Files

        fileName = CStr(file.name)
        extension = LCase$(fso.GetExtensionName(fileName))

        If extension = "xlsx" Or extension = "xls" Then

            ' Skip Excel temp files ($ prefix)
            If Left$(fileName, 2) <> "~$" Then

                ' Skip the macro workbook itself
                If Not PathsEqual(CStr(file.path), ThisWorkbook.FullName) Then
                    result.Add CStr(file.path)

                    Debug.Print "  Found: " & fileName
                End If

            End If

        End If

    Next file

    Set CollectInputFiles = SortCollection(result)

End Function


' ============================================================
' COLLECTION SORT (alphabetical, case-insensitive)
' ============================================================

Private Function SortCollection( _
    ByVal col As Collection) As Collection

    Dim arr() As String
    Dim i As Long, j As Long, tmp As String
    Dim result As New Collection

    If col.Count = 0 Then
        Set SortCollection = col
        Exit Function
    End If

    ReDim arr(1 To col.Count)

    For i = 1 To col.Count
        arr(i) = CStr(col.Item(i))
    Next i

    ' Simple alphabetical sort (sufficient for typical folder sizes)
    For i = 1 To col.Count - 1
        For j = i + 1 To col.Count
            If StrComp(arr(i), arr(j), vbTextCompare) > 0 Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i

    For i = 1 To UBound(arr)
        result.Add arr(i)
    Next i

    Set SortCollection = result

End Function


' ============================================================
' COPY RANGE — Determine the data range to copy
' Uses UsedRange; falls back to Find for very large ranges.
' ============================================================

Private Function GetCopyRange( _
    ByVal ws As Worksheet) As Range

    Dim ur As Range
    Dim lastCell As Range

    Set GetCopyRange = Nothing

    On Error GoTo ErrHandler

    Set ur = ws.UsedRange

    If ur Is Nothing Then
        Exit Function
    End If

    Debug.Print "  UsedRange: " & ur.Address(False, False) & _
        "  rows=" & ur.Rows.Count & "  cols=" & ur.Columns.Count

    ' --------------------------------------------------------
    ' Normal case: UsedRange is small enough
    ' --------------------------------------------------------

    If ur.Rows.Count <= LARGE_USERANGE_THRESHOLD Then
        Set GetCopyRange = ur
        Exit Function
    End If

    ' --------------------------------------------------------
    ' Large UsedRange: find the actual last data cell
    ' --------------------------------------------------------

    Debug.Print "  UsedRange > " & LARGE_USERANGE_THRESHOLD & " rows — searching last data cell..."

    Set lastCell = ws.Cells.Find( _
        What:="*", _
        After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)

    If lastCell Is Nothing Then
        Exit Function
    End If

    Debug.Print "  Last data cell: " & lastCell.Address(False, False)

    Set GetCopyRange = ws.Range( _
        ws.Cells(1, 1), _
        ws.Cells(lastCell.Row, lastCell.Column))

    Exit Function

ErrHandler:
    Set GetCopyRange = Nothing

End Function


' ============================================================
' FORMULA HANDLING
' Copy formulas that stay within the current sheet.
' External workbook references are converted to values
' (the values were already pasted in Stage 1).
' ============================================================

Private Sub CopyFormulasWithoutExternalLinks( _
    ByVal srcRange As Range, _
    ByVal tgt As Range)

    Dim formulaCells As Range
    Dim cell As Range
    Dim targetCell As Range
    Dim shiftedFormula As String

    Dim relativeRow As Long
    Dim relativeCol As Long

    On Error Resume Next

    Set formulaCells = srcRange.SpecialCells(xlCellTypeFormulas)

    On Error GoTo 0

    If formulaCells Is Nothing Then
        Debug.Print "  No formula cells."
        Exit Sub
    End If

    Debug.Print "  Formula cells: " & formulaCells.Count

    For Each cell In formulaCells.Cells

        relativeRow = cell.Row - srcRange.Row
        relativeCol = cell.Column - srcRange.Column

        If HasExternalReference(CStr(cell.Formula)) Then
            ' External reference (cross-sheet or cross-workbook) —
            ' keep the value already pasted in Stage 1
            ' (nothing else to do; value was set by xlPasteValuesAndNumberFormats)
        Else
            ' Internal formula — shift references by the offset to the
            ' target cell. Convert to R1C1-relative to the source cell,
            ' then assign via FormulaR1C1 so Excel re-anchors relative
            ' offsets at the destination (Excel's native copy semantics:
            ' absolute $ refs are preserved, relative refs shift).
            Set targetCell = tgt.Offset(relativeRow, relativeCol)

            shiftedFormula = ShiftFormulaR1C1(CStr(cell.Formula), cell)

            If Len(shiftedFormula) > 0 Then
                ' Normal path: assign converted R1C1 formula
                targetCell.FormulaR1C1 = shiftedFormula
            Else
                ' Fallback for long formulas (>255 chars):
                ' ConvertFormula has a length limit, so use
                ' Copy + PasteSpecial xlPasteFormulas which
                ' delegates offset shifting to Excel's engine.
                cell.Copy
                targetCell.PasteSpecial _
                    Paste:=xlPasteFormulas
                Application.CutCopyMode = False
            End If
        End If

    Next cell

End Sub


Private Function HasExternalReference( _
    ByVal formulaText As String) As Boolean

    ' --------------------------------------------------------
    ' External workbook reference: [Book.xlsx]Sheet1!A1
    ' --------------------------------------------------------

    If InStr(1, formulaText, "[", vbTextCompare) > 0 Then
        HasExternalReference = True
        Exit Function
    End If

    ' --------------------------------------------------------
    ' Cross-sheet reference: Sheet2!A1 or 'Sheet Name'!A1
    ' The '!' separates the sheet name from the cell/range.
    '
    ' Excludes Excel error constants which also contain '!':
    ' #REF!  #VALUE!  #DIV/0!  #NULL!  #NUM!
    ' --------------------------------------------------------

    Dim pos As Long
    pos = InStr(1, formulaText, "!", vbTextCompare)

    Do While pos > 0
        If pos > 1 And Mid$(formulaText, pos - 1, 1) = "#" Then
            ' Part of an error constant — skip
            pos = InStr(pos + 1, formulaText, "!", vbTextCompare)
        Else
            HasExternalReference = True
            Exit Function
        End If
    Loop

    HasExternalReference = False

End Function


' ============================================================
' FORMULA CONVERSION
' Convert an A1 formula string to R1C1 notation, anchored
' relative to the source cell. The caller assigns the result
' via .FormulaR1C1 on the target cell, which makes Excel
' re-anchor relative offsets at the destination position.
' This produces the same semantics as Excel's native
' copy/paste: absolute ($A$1) refs are preserved, relative
' (A1) refs are shifted by the positional delta.
' ============================================================

Private Function ShiftFormulaR1C1( _
    ByVal formulaText As String, _
    ByVal srcCell As Range) As String

    Dim result As Variant
    Dim longFormula As Boolean

    If Len(formulaText) > 250 Then longFormula = True

    ' --- Primary: Application.ConvertFormula A1 → R1C1 ---
    ' xlRelative forces $A$1 → R[..]C[..] (relative offsets).
    ' RelativeTo = srcCell so offsets are anchored to the source.
    On Error Resume Next

    If longFormula Then
        result = CVErr(xlErrValue)
    Else
        result = Application.ConvertFormula( _
            Formula:=formulaText, _
            FromReferenceStyle:=xlA1, _
            ToReferenceStyle:=xlR1C1, _
            ToAbsolute:=xlRelative, _
            RelativeTo:=srcCell)
    End If

    On Error GoTo 0

    If Not isError(result) Then
        ShiftFormulaR1C1 = CStr(result)
        Exit Function
    End If

    ' --- Fallback for long formulas (>255 chars): ---
    ' Use .FormulaR1C1 directly on the target cell, reading
    ' the source cell's R1C1 form (Excel handles the shift).
    ' Set a flag so the caller takes the Copy/PasteSpecial path.
    ShiftFormulaR1C1 = ""

End Function


' ============================================================
' PATH UTILITIES
' ============================================================

Private Function NormalizeFolderPath( _
    ByVal path As String) As String

    Dim p As String

    p = Trim$(path)

    If Right$(p, 1) <> "\" Then
        p = p & "\"
    End If

    NormalizeFolderPath = p

End Function


Private Function NormalizeFilePath( _
    ByVal path As String) As String

    NormalizeFilePath = Trim$(path)

End Function


Private Function PathsEqual( _
    ByVal path1 As String, _
    ByVal path2 As String) As Boolean

    Dim p1 As String
    Dim p2 As String

    p1 = Replace(Trim$(path1), "/", "\")
    p2 = Replace(Trim$(path2), "/", "\")

    PathsEqual = (StrComp(p1, p2, vbTextCompare) = 0)

End Function


Private Function FileExists( _
    ByVal filePath As String) As Boolean

    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)

End Function


' ============================================================
' OUTPUT FILE FORMAT
' ============================================================

Private Function GetFileFormat( _
    ByVal filePath As String) As XlFileFormat

    Dim ext As String

    ext = LCase$(Mid$(filePath, InStrRev(filePath, ".") + 1))

    Select Case ext
        Case "xlsx":  GetFileFormat = xlOpenXMLWorkbook
        Case "xls":   GetFileFormat = xlExcel8
        Case "xlsm":  GetFileFormat = xlOpenXMLWorkbookMacroEnabled
        Case Else:   GetFileFormat = xlOpenXMLWorkbook   ' default
    End Select

End Function


' ============================================================
' ROW / COLUMN / ZOOM PROPERTIES
'
' xlPasteFormats does NOT copy hidden rows, row heights,
' column widths, or zoom. These are handled separately.
' ============================================================

' ------------------------------------------------------------
' Hidden rows — copied for every sheet
' Spec 2.1: group consecutive hidden rows into ranges
' like "10:12,15:15" to avoid error 1004 from too complex
' comma-separated union ranges.
' ------------------------------------------------------------

Private Sub CopyHiddenRows( _
    ByVal srcWs As Worksheet, _
    ByVal outWs As Worksheet, _
    ByVal srcStartRow As Long, _
    ByVal rowCount As Long, _
    ByVal tgtStartRow As Long)

    Dim r As Long
    Dim gapStart As Long
    Dim inGap As Boolean

    gapStart = 0
    inGap = False

    For r = 0 To rowCount - 1
        If srcWs.Rows(srcStartRow + r).EntireRow.Hidden Then
            If Not inGap Then
                gapStart = r
                inGap = True
            End If
        Else
            If inGap Then
                outWs.Rows(tgtStartRow + gapStart).Resize( _
                    r - gapStart).EntireRow.Hidden = True
                inGap = False
            End If
        End If
    Next r

    ' Flush trailing block
    If inGap Then
        outWs.Rows(tgtStartRow + gapStart).Resize( _
            rowCount - gapStart).EntireRow.Hidden = True
    End If

End Sub


' ------------------------------------------------------------
' Row heights — copied for every sheet
' Spec 2.2: same height rows are batched into one assignment
' using Range(Rows(x), Rows(y)).RowHeight syntax.
' ------------------------------------------------------------

Private Sub CopyRowHeights( _
    ByVal srcWs As Worksheet, _
    ByVal outWs As Worksheet, _
    ByVal srcStartRow As Long, _
    ByVal rowCount As Long, _
    ByVal tgtStartRow As Long)

    Dim r As Long
    Dim startR As Long
    Dim curHeight As Double
    Dim nextHeight As Double

    r = 0

    Do While r < rowCount
        curHeight = srcWs.Rows(srcStartRow + r).RowHeight

        ' -1 = default height; skip (let target use its own default)
        If curHeight = -1 Then
            r = r + 1
        Else

            ' Find extent of rows with this same height
            startR = r
            Do While r < rowCount
                nextHeight = srcWs.Rows(srcStartRow + r).RowHeight
                If nextHeight = -1 Or nextHeight <> curHeight Then Exit Do
                r = r + 1
            Loop

            ' Batch-set contiguous rows (same height)
            outWs.Range( _
                outWs.Rows(tgtStartRow + startR), _
                outWs.Rows(tgtStartRow + r - 1)).RowHeight = curHeight
        End If
    Loop

End Sub


' ------------------------------------------------------------
' Column widths — copied from the first data sheet only
' Spec 3.2: after the first data sheet, column widths are
' not updated on subsequent sheets.
' ------------------------------------------------------------

Private Sub CopyColumnWidths( _
    ByVal srcWs As Worksheet, _
    ByVal outWs As Worksheet, _
    ByVal srcStartCol As Long, _
    ByVal colCount As Long, _
    ByVal tgtStartCol As Long)

    Dim c As Long
    Dim cw As Double

    For c = 0 To colCount - 1
        cw = srcWs.Columns(srcStartCol + c).ColumnWidth

        ' -1 = default width; skip (target uses its own default)
        If cw >= 0 Then
            outWs.Columns(tgtStartCol + c).ColumnWidth = cw
        End If
    Next c

End Sub


' ------------------------------------------------------------
' Zoom — copied from the first data sheet only
' Spec 3.3: Zoom is a Window property, so source sheet
' activation is required to read it. ScreenUpdating is
' already False, so the activation is invisible.
' ------------------------------------------------------------

Private Sub CopyZoom( _
    ByVal srcWb As Workbook, _
    ByVal srcWs As Worksheet, _
    ByVal outWb As Workbook)

    Dim zoomVal As Variant

    srcWs.Activate
    zoomVal = srcWb.Windows(1).Zoom

    outWb.Windows(1).Zoom = zoomVal

    ThisWorkbook.Activate

End Sub


' ------------------------------------------------------------
' Column hidden state — copied from first data sheet only
' Only copies the .Hidden property (visible state), not
' column widths or other formatting.
' ------------------------------------------------------------

Private Sub CopyColumnHidden( _
    ByVal srcWs As Worksheet, _
    ByVal outWs As Worksheet, _
    ByVal srcStartCol As Long, _
    ByVal colCount As Long, _
    ByVal tgtStartCol As Long)

    Dim c As Long

    For c = 0 To colCount - 1
        If srcWs.Columns(srcStartCol + c).Hidden Then
            outWs.Columns(tgtStartCol + c).Hidden = True
        End If
    Next c

End Sub
