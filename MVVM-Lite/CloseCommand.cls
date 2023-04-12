VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CloseCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICommand

Private Function ICommand_CanExecute(ByVal Parameter As Object) As Boolean
    ICommand_CanExecute = TypeOf Parameter Is MSForms.UserForm
End Function

Private Sub ICommand_Execute(ByVal Parameter As Object)
    Dim Form As OrderForm
    Set Form = Parameter
    Form.Hide
End Sub
