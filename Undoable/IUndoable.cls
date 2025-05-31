VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IUndoable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Undoable")
'@Interface
Option Explicit

'@Description("Undoes a previously performed action")
Public Sub Undo()
Attribute Undo.VB_Description = "Undoes a previously performed action"
End Sub

'@Description("Redoes a previously undone action")
Public Sub Redo()
Attribute Redo.VB_Description = "Redoes a previously undone action"
End Sub

'@Description("Describes the undoable action")
Public Property Get Description() As String
Attribute Description.VB_Description = "Describes the undoable action"
End Property
