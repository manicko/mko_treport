Attribute VB_Name = "modTreportState"
Option Explicit

Public Type AppState
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    DisplayAlerts As Boolean
    Calculation As XlCalculation
    AskToUpdateLinks As Boolean
    StatusBar As Variant
End Type

Public Function CaptureState() As AppState
    Dim s As AppState
    s.ScreenUpdating = Application.ScreenUpdating
    s.EnableEvents = Application.EnableEvents
    s.DisplayAlerts = Application.DisplayAlerts
    s.Calculation = Application.Calculation
    s.AskToUpdateLinks = Application.AskToUpdateLinks
    s.StatusBar = Application.StatusBar
    CaptureState = s
End Function

Public Sub ApplyFastState(ByRef s As AppState)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    Application.AskToUpdateLinks = False
    On Error GoTo 0
End Sub

Public Sub RestoreState(ByRef s As AppState)
    On Error Resume Next
    Application.ScreenUpdating = s.ScreenUpdating
    Application.EnableEvents = s.EnableEvents
    Application.DisplayAlerts = s.DisplayAlerts
    Application.Calculation = s.Calculation
    Application.AskToUpdateLinks = s.AskToUpdateLinks
    Application.StatusBar = False
    Application.CutCopyMode = False
    On Error GoTo 0
End Sub









