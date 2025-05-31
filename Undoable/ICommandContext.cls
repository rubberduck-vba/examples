VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICommandContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Undoable.Commands")
'@Interface
Option Explicit

'@Description("True if the model is valid in its current state")
Public Function IsValid() As Boolean
Attribute IsValid.VB_Description = "True if the model is valid in its current state"
End Function
