VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClearLineItemsCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICommand

Private Function ICommand_CanExecute(ByVal Parameter As Object) As Boolean
    ICommand_CanExecute = TypeOf Parameter Is OrderHeaderModel
End Function

Private Sub ICommand_Execute(ByVal Parameter As Object)
    If Not TypeOf Parameter Is OrderHeaderModel Then Err.Raise 5
    
    Dim Model As OrderHeaderModel
    Set Model = Parameter
        
    Model.ClearLineItems
    
End Sub

