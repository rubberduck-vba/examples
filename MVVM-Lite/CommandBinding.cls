VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents UI As MSForms.CommandButton
Attribute UI.VB_VarHelpID = -1
Private Source As ICommand
Private Args As Object

Public Sub Initialize(ByVal Button As MSForms.CommandButton, ByVal Command As ICommand, ByVal Parameter As Object)
    Set UI = Button
    Set Source = Command
    Set Args = Parameter
End Sub

Private Sub UI_Click()
    Source.Execute Args
End Sub
