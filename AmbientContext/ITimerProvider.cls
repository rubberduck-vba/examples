VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITimerProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A service that ensures all clients get the same IAppTimer instance."
'@Folder "AmbientContext.Abstract"
'@ModuleDescription "A service that ensures all clients get the same IAppTimer instance."
'@Interface
Option Explicit

'@Description "Gets an IAppTimer instance."
Public Property Get Timer() As IAppTimer
Attribute Timer.VB_Description = "Gets an IAppTimer instance."
End Property
