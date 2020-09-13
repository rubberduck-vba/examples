VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CancelCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A command that closes (hides) a cancellable View in a cancelled state."
'@Folder MVVM.Infrastructure.Commands
'@ModuleDescription "A command that closes (hides) a cancellable View in a cancelled state."
'@PredeclaredId
'@Exposed
Option Explicit
Implements ICommand

Private Type TState
    View As ICancellable
End Type

Private this As TState

'@Description "Creates a new instance of this command."
Public Function Create(ByVal View As ICancellable) As ICommand
Attribute Create.VB_Description = "Creates a new instance of this command."
    Dim result As CancelCommand
    Set result = New CancelCommand
    Set result.View = View
    Set Create = result
End Function

Public Property Get View() As ICancellable
    Set View = this.View
End Property

Public Property Set View(ByVal RHS As ICancellable)
    GuardClauses.GuardDoubleInitialization this.View, TypeName(Me)
    Set this.View = RHS
End Property

Private Function ICommand_CanExecute(ByVal Context As Object) As Boolean
    ICommand_CanExecute = True
End Function

Private Property Get ICommand_Description() As String
    ICommand_Description = "Cancel pending changes and close."
End Property

Private Sub ICommand_Execute(ByVal Context As Object)
    this.View.OnCancel
End Sub
