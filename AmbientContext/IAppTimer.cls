VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAppTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "AmbientContext.Abstract"
'@Interface
Option Explicit

'@Description "Schedules the execution of a callback. Invoke from a non-volatile UDF."
Public Sub ExecuteMacroAsync()
Attribute ExecuteMacroAsync.VB_Description = "Schedules the execution of a callback. Invoke from a non-volatile UDF."
End Sub

'@Description "Schedules the execution of a side-effecting macro. Invoked from the scheduled Win32 callback."
Public Sub OnCallback()
Attribute OnCallback.VB_Description = "Schedules the execution of a side-effecting macro. Invoked from the scheduled Win32 callback."
End Sub
