VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimerProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "AmbientContext"
Option Explicit
Implements ITimerProvider

Private Const MacroName As String = "Execute"

Private Property Get ITimerProvider_Timer() As IAppTimer
    Static Instance As AppTimer
    If Instance Is Nothing Then
        Set Instance = New AppTimer
        Instance.MacroName = MacroName
    End If
    Set ITimerProvider_Timer = Instance
End Property
