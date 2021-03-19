Attribute VB_Name = "Macros"
'@Folder "UDF"
Option Private Module
Option Explicit

'@Description "Invoked from the Win32 timer callback."
Public Sub ExecuteMacroAsync()
Attribute ExecuteMacroAsync.VB_Description = "Invoked from the Win32 timer callback."
    'get the macro to run from the AppContext default instance:
    AppContext.Timer.OnCallback
End Sub

'Description "The actual side-effecting code, indirectly invoked from AppTimer."
Public Sub Execute()
    With AppContext.Current
        .Target.Value = .Property("Test1")
        .Target.Offset(1).Value = .Property("Test2")
        .Clear
    End With
End Sub
