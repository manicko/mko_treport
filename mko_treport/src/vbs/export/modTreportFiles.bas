Attribute VB_Name = "modTreportFiles"
Option Explicit

' treport rev.5 - file resolution and workbook open/close helpers (B6)

Public Function ResolveDataFilePath(ByVal inputFolder As String, ByVal baseName As String) As String
    Dim fso As Object, exts, i As Long, p As String
    Dim foundPath As String, foundCount As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    exts = Array(".xlsx", ".xlsm", ".xlsb") ' precedence D2
    foundCount = 0
    foundPath = vbNullString
    For i = LBound(exts) To UBound(exts)
        p = inputFolder & baseName & exts(i)
        If fso.FileExists(p) Then
            foundCount = foundCount + 1
            If Len(foundPath) = 0 Then foundPath = p
        End If
    Next i
    If foundCount = 0 Then
        modTreportConfig.Fail 2201, "files/data_file='" & baseName & "',input_path='" & inputFolder & "'", _
            "No matching data file found (checked .xlsx,.xlsm,.xlsb)."
    ElseIf foundCount > 1 Then
        modTreportConfig.Fail 2202, "files/data_file='" & baseName & "',input_path='" & inputFolder & "'", _
            "Multiple matching data files found for base name (ambiguous)."
    End If
    ResolveDataFilePath = foundPath
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
