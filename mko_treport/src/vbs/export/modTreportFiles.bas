Attribute VB_Name = "modTreportFiles"
Option Explicit

' treport - file resolution and workbook open/close helpers (D2 / FR-FILE-1 / FR-FILE-2)
' D2 RC1: resolve data_files.data_file (base name) in general.input_path directory only
' (non-recursive). Allowed extensions: .xlsx,.xlsm,.xlsb (fixed set/scan order, not a
' precedence tie-breaker). Exactly one candidate -> select; zero -> no match fail;
' >1 candidate for same base name across allowed extensions -> ambiguous fail-fast.

Private Const FILE_EXTS As String = ".xlsx,.xlsm,.xlsb"

Public Function ResolveDataFilePath(ByVal inputFolder As String, ByVal baseName As String) As String
    Dim fso As Object
    Dim exts As Variant
    Dim matches As Collection
    Dim ctx As String
    Dim searchedPath As String
    Dim i As Long
    Dim candList As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    exts = Split(FILE_EXTS, ",")
    searchedPath = fso.GetAbsolutePathName(inputFolder)
    ctx = "files/data_file='" & baseName & "',input_path='" & searchedPath & "'"

    If Not fso.FolderExists(searchedPath) Then
        modTreportConfig.Fail 2201, ctx, _
            "Cause: input_path folder not found. Searched path: '" & searchedPath & _
            "' for data_file base name '" & baseName & "' (allowed extensions " & FILE_EXTS & ")."
    End If

    Set matches = CollectMatchingFiles(fso, searchedPath, baseName, exts)

    If matches.Count = 0 Then
        modTreportConfig.Fail 2201, ctx, _
            "Cause: no matching data file found. Searched path: '" & searchedPath & _
            "' (non-recursive) for data_file base name '" & baseName & _
            "' with allowed extensions " & FILE_EXTS & "."
    End If

    If matches.Count > 1 Then
        candList = ""
        For i = 1 To matches.Count
            If Len(candList) > 0 Then candList = candList & ", "
            candList = candList & "'" & CStr(matches(i)) & "'"
        Next i
        modTreportConfig.Fail 2202, ctx & ",matches=" & CStr(matches.Count) & ",candidates=" & candList, _
            "Cause: ambiguous data file resolution (>1 matching candidate for same base name). " & _
            "Searched path: '" & searchedPath & "' (non-recursive) for data_file base name '" & baseName & _
            "' with allowed extensions " & FILE_EXTS & " (scan order only, not a precedence tie-breaker). " & _
            "Candidates: " & candList & "."
    End If

    ResolveDataFilePath = CStr(matches(1))
End Function

Private Function CollectMatchingFiles(ByVal fso As Object, ByVal folderPath As String, ByVal baseName As String, ByRef exts As Variant) As Collection
    Dim col As New Collection
    Dim i As Long
    Dim ext As String
    Dim candidate As String

    ' Deterministic scan in fixed allowed-extension order (order is not a precedence tie-breaker).
    For i = LBound(exts) To UBound(exts)
        ext = CStr(exts(i))
        candidate = fso.BuildPath(folderPath, baseName & ext)
        If fso.FileExists(candidate) Then
            col.Add candidate
        End If
    Next i

    Set CollectMatchingFiles = col
End Function

Public Function OpenWorkbookReadOnly(ByVal path As String) As Workbook
    Dim wb As Workbook
    On Error GoTo EH
    Set wb = Application.Workbooks.Open(fileName:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    Set OpenWorkbookReadOnly = wb
    Exit Function
EH:
    modTreportConfig.Fail 2203, "files/open_readonly='" & path & "'", "Unable to open workbook read-only: " & Err.Description
End Function

Public Function OpenWorkbookWritable(ByVal path As String) As Workbook
    Dim wb As Workbook
    On Error GoTo EH
    Set wb = Application.Workbooks.Open(fileName:=path, UpdateLinks:=0, ReadOnly:=False, AddToMru:=False)
    On Error GoTo 0
    If wb.ReadOnly Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
        modTreportConfig.Fail 2204, "files/open_writable='" & path & "'", "Workbook opened as read-only/locked; Refresh_data requires writable access."
    End If
    Set OpenWorkbookWritable = wb
    Exit Function
EH:
    modTreportConfig.Fail 2205, "files/open_writable='" & path & "'", "Unable to open workbook for refresh: " & Err.Description
End Function

Public Sub CloseNoSave(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Public Sub SaveAndClose(ByVal wb As Workbook, ByVal path As String)
    If wb Is Nothing Then Exit Sub
    On Error GoTo EH
    wb.Save
    wb.Close SaveChanges:=False
    Exit Sub
EH:
    modTreportConfig.Fail 2206, "files/save='" & path & "'", "Unable to save refreshed workbook: " & Err.Description
End Sub
