Attribute VB_Name = "modTreportFiles"
Option Explicit

' treport rev.5 - file resolution and workbook open/close helpers (B6, D1, D2, FR-FILE-1/FR-FILE-2)

Public Function ResolveDataFilePath(ByVal inputFolder As String, ByVal baseName As String) As String
    Dim fso As Object
    Dim exts As Variant
    Dim allMatches As Collection
    Dim ctx As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    exts = Array(".xlsx", ".xlsm", ".xlsb")
    ctx = "files/data_file='" & baseName & "',input_path='" & inputFolder & "'"

    Set allMatches = CollectMatchingFiles(fso, inputFolder, baseName, exts)
    If allMatches.Count = 0 Then
        modTreportConfig.Fail 2201, ctx, _
            "No matching data file found (checked .xlsx,.xlsm,.xlsb in input_path and subfolders)."
    End If

    If allMatches.Count <> 1 Then
        modTreportConfig.Fail 2202, ctx & ",matches=" & allMatches.Count, _
            "Multiple matching data files found for base name; expected exactly one match within input_path/subfolders."
    End If

    ResolveDataFilePath = CStr(allMatches(1))
End Function

Private Function CollectMatchingFiles(ByVal fso As Object, ByVal folderPath As String, ByVal baseName As String, ByRef exts As Variant) As Collection
    Dim col As New Collection
    Dim fld As Object
    If Not fso.FolderExists(folderPath) Then
        Set CollectMatchingFiles = col
        Exit Function
    End If
    Set fld = fso.GetFolder(folderPath)
    CollectMatchingFilesInFolder fso, fld, baseName, exts, col
    Set CollectMatchingFiles = col
End Function

Private Sub CollectMatchingFilesInFolder(ByVal fso As Object, ByVal fld As Object, ByVal baseName As String, ByRef exts As Variant, ByRef col As Collection)
    Dim fil As Object, sf As Object
    Dim nameNoExt As String, ext As String, full As String
    Dim i As Long
    For Each fil In fld.Files
        nameNoExt = fso.GetBaseName(fil.Name)
        If StrComp(nameNoExt, baseName, vbTextCompare) = 0 Then
            ext = "." & LCase$(fso.GetExtensionName(fil.Name))
            For i = LBound(exts) To UBound(exts)
                If StrComp(ext, CStr(exts(i)), vbTextCompare) = 0 Then
                    full = fil.path
                    col.Add full
                    Exit For
                End If
            Next i
        End If
    Next fil
    For Each sf In fld.SubFolders
        ' recurse into subfolders (e.g. monthly/weekly/base)
        CollectMatchingFilesInFolder fso, sf, baseName, exts, col
    Next sf
End Sub

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
