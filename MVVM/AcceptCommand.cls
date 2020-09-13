VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AcceptCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A command that closes (hides) a View."
'@Exposed
'@Folder MVVM.Infrastructure.Commands
'@ModuleDescription "A command that closes (hides) a View."
'@PredeclaredId
Option Explicit
Implements ICommand

Private Type TState
    View As IView
End Type

Private this As TState

'@Description "Creates a new instance of this command."
Public Function Create(ByVal View As IView) As ICommand
Attribute Create.VB_Description = "Creates a new instance of this command."
    Dim result As AcceptCommand
    Set result = New AcceptCommand
    Set result.View = View
    Set Create = result
End Function

Public Property Get View() As IView
    Set View = this.View
End Property

Public Property Set View(ByVal RHS As IView)
    GuardClauses.GuardDoubleInitialization this.View, TypeName(Me)
    Set this.View = RHS
End Property

Private Function ICommand_CanExecute(ByVal Context As Object) As Boolean
    Dim ViewModel As IViewModel
    If TypeOf Context Is IViewModel Then
        Set ViewModel = Context
        If Not ViewModel.Validation Is Nothing Then
            ICommand_CanExecute = ViewModel.Validation.IsValid
            Exit Function
        End If
    End If
    ICommand_CanExecute = True
End Function

Private Property Get ICommand_Description() As String
    ICommand_Description = "Accept changes and close."
End Property

Private Sub ICommand_Execute(ByVal Context As Object)
    this.View.Hide
End Sub
