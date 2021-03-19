VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "AmbientContext"
Option Explicit
Implements IAppTimer
'based on https://stackoverflow.com/a/9493006/1188513
Private Declare Function SetTimer Lib "user32" (ByVal HWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal HWnd As Long, ByVal nIDEvent As Long) As Long

Private Type TMacro
    TimerId As Long
    TimerTime As Date
    MacroName As String
End Type

Private This As TMacro

'@Ignore ProcedureNotUsed
Public Property Get MacroName() As String
    MacroName = This.MacroName
End Property

Public Property Let MacroName(ByVal Value As String)
    This.MacroName = Value
End Property

Private Sub TerminateTimer()
    If This.TimerId = 0 Then Exit Sub
    On Error Resume Next
    KillTimer 0&, This.TimerId
    On Error GoTo 0
    This.TimerId = 0
End Sub

Private Sub ExecuteInternal()

   If This.TimerTime <> 0 Then
      On Error Resume Next
      Application.OnTime This.TimerTime, This.MacroName, Schedule:=False
      On Error GoTo 0
   End If
   
   This.TimerTime = DateTime.Now
   
   On Error Resume Next
   Application.OnTime This.TimerTime, This.MacroName
   On Error GoTo 0

End Sub

Private Sub Class_Initialize()
    This.MacroName = "Execute"
End Sub

Private Sub Class_Terminate()
    TerminateTimer
End Sub

Private Sub IAppTimer_ExecuteMacroAsync()
    TerminateTimer
    This.TimerId = SetTimer(0&, 0&, 1, AddressOf Macros.ExecuteMacroAsync)
End Sub

Private Sub IAppTimer_OnCallback()
    TerminateTimer
    ExecuteInternal
End Sub
