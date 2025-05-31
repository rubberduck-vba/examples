VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Undoable.Commands"
'@Interface
Option Explicit

'@Description("Returns True if the command can be executed given the provided context")
Public Function CanExecute(ByVal Context As Object) As Boolean
Attribute CanExecute.VB_Description = "Returns True if the command can be executed given the provided context"
End Function

'@Description("Executes an action given a context")
Public Sub Execute(ByVal Context As Object)
Attribute Execute.VB_Description = "Executes an action given a context"
End Sub
