Attribute VB_Name = "modTreportRefresh"
Option Explicit

' treport rev.5 - Refresh_data workflow (B7, D1, FR-REF)

Public Sub RefreshWorkbookData(ByVal wb As Workbook, ByVal ctx As String)
    Dim conn As WorkbookConnection
    Dim qt As QueryTable
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim pt As PivotTable
    Dim pc As PivotCache
    Dim t As Single

    On Error GoTo EH

    ' Force sync where possible
    For Each conn In wb.Connections
        On Error Resume Next
        conn.OLEDBConnection.BackgroundQuery = False
        conn.ODBCConnection.BackgroundQuery = False
        On Error GoTo EH
    Next conn

    For Each ws In wb.Worksheets
        For Each qt In ws.QueryTables
            On Error Resume Next
            qt.BackgroundQuery = False
            On Error GoTo EH
        Next qt
        On Error Resume Next
        For Each lo In ws.ListObjects
            If Not lo.QueryTable Is Nothing Then lo.QueryTable.BackgroundQuery = False
        Next lo
        On Error GoTo EH
    Next ws

    For Each pc In wb.PivotCaches
        On Error Resume Next
        pc.BackgroundQuery = False
        On Error GoTo EH
    Next pc

    wb.RefreshAll
    On Error Resume Next
    Application.CalculateUntilAsyncQueriesDone
    On Error GoTo EH
    DoEvents

    ' wait for connections/queries with timeout
    t = Timer
    Do
        Dim busy As Boolean
        busy = False
        On Error Resume Next
        For Each conn In wb.Connections
            If conn.OLEDBConnection.Refreshing Then busy = True
            If conn.ODBCConnection.Refreshing Then busy = True
        Next conn
        For Each ws In wb.Worksheets
            For Each qt In ws.QueryTables
                If qt.Refreshing Then busy = True
            Next qt
            For Each lo In ws.ListObjects
                If Not lo.QueryTable Is Nothing Then
                    If lo.QueryTable.Refreshing Then busy = True
                End If
            Next lo
        Next ws
        On Error GoTo EH
        If Not busy Then Exit Do
        DoEvents
        If Timer - t > 300 Then modTreportConfig.Fail 2501, ctx, "Refresh timed out waiting for data connections/queries to complete."
    Loop

    On Error Resume Next
    Application.CalculateUntilAsyncQueriesDone
    On Error GoTo EH

    ' refresh pivot caches after source refresh
    For Each pc In wb.PivotCaches
        On Error Resume Next
        pc.MissingItemsLimit = xlMissingItemsNone
        pc.Refresh
        If Err.Number <> 0 Then
            Err.Clear
            pc.Refresh
        End If
        On Error GoTo EH
    Next pc
    For Each ws In wb.Worksheets
        For Each pt In ws.PivotTables
            On Error Resume Next
            pt.RefreshTable
            On Error GoTo EH
        Next pt
    Next ws

    Exit Sub
EH:
    modTreportConfig.Fail 2502, ctx, "Refresh failed: " & Err.Description
End Sub










