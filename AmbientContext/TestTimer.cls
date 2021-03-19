VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Tests.Stubs"
Option Explicit
Implements IAppTimer

Private Type TState
    ExecuteMacroAsyncInvokes As Long
    OnCallbackInvokes As Long
End Type

Private This As TState

Public Property Get ExecuteMacroAsyncInvokes() As Long
    ExecuteMacroAsyncInvokes = This.ExecuteMacroAsyncInvokes
End Property

Public Property Get OnCallbackInvokes() As Long
    OnCallbackInvokes = This.OnCallbackInvokes
End Property

Private Sub IAppTimer_ExecuteMacroAsync()
    This.ExecuteMacroAsyncInvokes = This.ExecuteMacroAsyncInvokes + 1
End Sub

Private Sub IAppTimer_OnCallback()
    This.OnCallbackInvokes = This.OnCallbackInvokes + 1
End Sub
