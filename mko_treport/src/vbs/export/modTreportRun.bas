Attribute VB_Name = "modTreportRun"
Option Explicit

' ============================================================
' treport rev.5 run entry points (B1 scaffold)
'
' Required UI bindings (run sheet):
'   - Button/control "Refresh_data"  -> modTreportRun.Refresh_data
'   - Button/control "Create_report" -> modTreportRun.Create_report
'
' Legacy RunXlconcat/xlconcat_macro workflow must not be bound
' to any run-sheet control.
'
' Full business behavior is implemented in later blocks (B4-B12).
' These stubs fail fast with an explicit, non-destructive message.
' ============================================================

Private Const MODULE_NAME As String = "modTreportRun"
Private Const NOT_IMPLEMENTED_MSG As String = _
    " is not yet implemented for rev.5 (B1 control surface only). " & _
    "No data was refreshed, no report was generated, and no legacy xlconcat workflow was executed."

Private Sub ReportNotImplemented(ByVal entryPointName As String)
    MsgBox entryPointName & NOT_IMPLEMENTED_MSG, vbExclamation, "treport"
End Sub

Public Sub Refresh_data()
    ReportNotImplemented "Refresh_data"
End Sub

Public Sub Create_report()
    ReportNotImplemented "Create_report"
End Sub
