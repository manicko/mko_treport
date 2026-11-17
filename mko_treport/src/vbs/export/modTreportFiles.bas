Attribute VB_Name = "modTreportFiles"
Option Explicit

' treport - file resolution and workbook open/close helpers (D2 / FR-FILE-1 / FR-FILE-2)
' D2 (operational): resolve data_files.data_file (base name) within the general.input_path
' subtree (input_path and all descendant subfolders, recursive). Allowed extensions:
' .xlsx,.xlsm,.xlsb (fixed set/scan order, not a precedence tie-breaker). Exactly one
' candidate across the full searched scope -> select; zero -> no match fail-fast;
' >1 candidate for same base name across allowed extensions/locations -> ambiguous fail-fast
' with full candidate paths listed. Resolution is deterministic and repeatable.

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
            "' (including subfolders) for data_file base name '" & baseName & "' (allowed extensions " & FILE_EXTS & ")."
    End If

    Set matches = CollectMatchingFiles(fso, searchedPath, baseName, exts)
    Set matches = SortStringCollection(matches)

    If matches.Count = 0 Then
        modTreportConfig.Fail 2201, ctx, _
            "Cause: no matching data file found. Searched path: '" & searchedPath & _
            "' (including subfolders) for data_file base name '" & baseName & _
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
            "Searched path: '" & searchedPath & "' (including subfolders) for data_file base name '" & baseName & _
            "' with allowed extensions " & FILE_EXTS & " (scan order only, not a precedence tie-breaker). " & _
            "Candidates: " & candList & "."
    End If

    ResolveDataFilePath = CStr(matches(1))
End Function

Private Function CollectMatchingFiles(ByVal fso As Object, ByVal folderPath As String, ByVal baseName As String, ByRef exts As Variant) As Collection
    Dim col As New Collection
    CollectMatchingFilesInFolder fso, folderPath, baseName, exts, col
    Set CollectMatchingFiles = col
End Function

Private Sub CollectMatchingFilesInFolder(ByVal fso As Object, ByVal folderPath As String, ByVal baseName As String, ByRef exts As Variant, ByRef col As Collection)
    Dim fld As Object
    Dim subFld As Object
    Dim fil As Object
    Dim i As Long
    Dim ext As String
    Dim targetBase As String
    Dim fileBase As String
    Dim fullPath As String

    Set fld = fso.GetFolder(folderPath)
    targetBase = LCase$(baseName)

    ' Deterministic scan: fixed allowed-extension order per folder level (order is not a
    ' precedence tie-breaker when multiple candidates exist across scope).
    For i = LBound(exts) To UBound(exts)
        ext = LCase$(CStr(exts(i)))
        For Each fil In fld.Files
            fullPath = CStr(fil.path)
            If LCase$(fso.GetExtensionName(fullPath)) = Mid$(ext, 2) Then
                fileBase = fso.GetBaseName(fullPath)
                If StrComp(fileBase, baseName, vbTextCompare) = 0 Or LCase$(fileBase) = targetBase Then
                    col.Add fullPath
                End If
            End If
        Next fil
    Next i

    ' Recurse into all descendant subfolders (stable traversal by collected subfolder paths).
    Dim subPaths As Collection
    Set subPaths = New Collection
    For Each subFld In fld.SubFolders
        subPaths.Add CStr(subFld.path)
    Next subFld
    Set subPaths = SortStringCollection(subPaths)
    For i = 1 To subPaths.Count
        CollectMatchingFilesInFolder fso, CStr(subPaths(i)), baseName, exts, col
    Next i
End Sub

Private Function SortStringCollection(ByVal col As Collection) As Collection
    Dim sorted As New Collection
    Dim arr() As String
    Dim i As Long
    Dim j As Long
    Dim tmp As String

    If col Is Nothing Then
        Set SortStringCollection = sorted
        Exit Function
    End If
    If col.Count <= 1 Then
        For i = 1 To col.Count
            sorted.Add CStr(col(i))
        Next i
        Set SortStringCollection = sorted
        Exit Function
    End If

    ReDim arr(1 To col.Count)
    For i = 1 To col.Count
        arr(i) = CStr(col(i))
    Next i
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(arr(i), arr(j), vbBinaryCompare) > 0 Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i
    For i = 1 To UBound(arr)
        sorted.Add arr(i)
    Next i
    Set SortStringCollection = sorted
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












