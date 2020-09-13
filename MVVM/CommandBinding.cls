VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for binding a command to a UI element."
'@Folder MVVM.Infrastructure.Bindings
'@ModuleDescription "An object responsible for binding a command to a UI element."
'@PredeclaredId
'@Exposed
Implements ICommandBinding
Option Explicit

Private Type TCommandBinding
    ViewModel As Object
    Target As Object
    Command As ICommand
End Type

Private WithEvents CommandButtonEvents As MSForms.CommandButton
Attribute CommandButtonEvents.VB_VarHelpID = -1
Private WithEvents CheckBoxEvents As MSForms.CheckBox
Attribute CheckBoxEvents.VB_VarHelpID = -1
Private WithEvents ImageEvents As MSForms.Image
Attribute ImageEvents.VB_VarHelpID = -1
Private WithEvents LabelEvents As MSForms.Label
Attribute LabelEvents.VB_VarHelpID = -1

Private this As TCommandBinding

'@Description "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.CommandButton target."
Public Function ForCommandButton(ByVal Target As MSForms.CommandButton, ByVal Command As ICommand, ByVal ViewModel As Object) As ICommandBinding
Attribute ForCommandButton.VB_Description = "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.CommandButton target."
    GuardClauses.GuardNonDefaultInstance Me, CommandBinding
    Set ForCommandButton = Create(Target, Command, ViewModel)
End Function

'@Description "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.CheckBox target."
Public Function ForCheckBox(ByVal Target As MSForms.CheckBox, ByVal Command As ICommand, ByVal ViewModel As Object) As ICommandBinding
Attribute ForCheckBox.VB_Description = "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.CheckBox target."
    GuardClauses.GuardNonDefaultInstance Me, CommandBinding
    Set ForCheckBox = Create(Target, Command, ViewModel)
End Function

'@Description "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.Image target."
Public Function ForImage(ByVal Target As MSForms.Image, ByVal Command As ICommand, ByVal ViewModel As Object) As ICommandBinding
Attribute ForImage.VB_Description = "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.Image target."
    GuardClauses.GuardNonDefaultInstance Me, CommandBinding
    Set ForImage = Create(Target, Command, ViewModel)
End Function

'@Description "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.Label target."
Public Function ForLabel(ByVal Target As MSForms.Label, ByVal Command As ICommand, ByVal ViewModel As Object) As ICommandBinding
Attribute ForLabel.VB_Description = "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified MSForms.Label target."
    GuardClauses.GuardNonDefaultInstance Me, CommandBinding
    Set ForLabel = Create(Target, Command, ViewModel)
End Function

'@Description "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified Target."
Public Function Create(ByVal Target As Object, ByVal Command As ICommand, ByVal ViewModel As Object) As ICommandBinding
Attribute Create.VB_Description = "Creates and returns an ICommandBinding implementation binding the specified ICommand to the specified Target."
    GuardClauses.GuardNonDefaultInstance Me, CommandBinding
    
    Dim result As CommandBinding
    Set result = New CommandBinding
    
    Set result.ViewModel = ViewModel
    Set result.Target = Target
    Set result.Command = Command
    
    Set Create = result
    
End Function

Public Property Get ViewModel() As Object
    Set ViewModel = this.ViewModel
End Property

Public Property Set ViewModel(ByVal RHS As Object)
    Set this.ViewModel = RHS
End Property

Public Property Get Target() As Object
    Set Target = this.Target
End Property

Public Property Set Target(ByVal RHS As Object)

    GuardClauses.GuardDoubleInitialization this.Target, TypeName(Me)
    Set this.Target = RHS
    
    Select Case True
        Case TypeOf RHS Is MSForms.CommandButton
            Set CommandButtonEvents = RHS
        
        Case TypeOf RHS Is MSForms.CheckBox
            Set CheckBoxEvents = RHS
        
        Case TypeOf RHS Is MSForms.Image
            Set ImageEvents = RHS
            
        Case TypeOf RHS Is MSForms.Label
            Set LabelEvents = RHS
    
        Case Else
            GuardClauses.GuardExpression _
                Throw:=True, _
                Source:=TypeName(Me), _
                Message:="Type '" & TypeName(RHS) & "' does not support command bindings at the moment."
    End Select

End Property

Public Property Get Command() As ICommand
    Set Command = this.Command
End Property

Public Property Set Command(ByVal RHS As ICommand)
    Set this.Command = RHS
    If Not RHS Is Nothing And Not this.Target Is Nothing Then
        this.Target.ControlTipText = RHS.Description
    End If
End Property

Private Property Get AsInterface() As ICommandBinding
    Set AsInterface = Me
End Property

Private Sub OnExecute()
    If Not this.Command Is Nothing Then
        this.Command.Execute this.ViewModel
    Else
        Debug.Print "BUG in " & TypeName(Me) & ": Command is 'Nothing', cannot execute."
        Debug.Assert False ' should not happen, break here if it does.
    End If
End Sub

Private Property Get ICommandBinding_Target() As Object
    Set ICommandBinding_Target = this.Target
End Property

Private Property Get ICommandBinding_Command() As ICommand
    Set ICommandBinding_Command = this.Command
End Property

Private Sub ICommandBinding_EvaluateCanExecute(ByVal Context As Object)
    EvaluateCanExecute Context
End Sub

Private Sub EvaluateCanExecute(ByVal Source As Object)
    If this.Target Is Nothing Then Exit Sub
    If this.Command Is Nothing Then
        this.Target.Enabled = False
    Else
        this.Target.Enabled = this.Command.CanExecute(Source)
    End If
End Sub

Private Sub CheckBoxEvents_Click()
    GuardClauses.GuardExpression Not TypeOf this.Target Is MSForms.CheckBox, TypeName(Me)
    OnExecute
End Sub

Private Sub CommandButtonEvents_Click()
    GuardClauses.GuardExpression Not TypeOf this.Target Is MSForms.CommandButton, TypeName(Me)
    OnExecute
End Sub

Private Sub ImageEvents_Click()
    GuardClauses.GuardExpression Not TypeOf this.Target Is MSForms.Image, TypeName(Me)
    OnExecute
End Sub

Private Sub LabelEvents_Click()
    GuardClauses.GuardExpression Not TypeOf this.Target Is MSForms.Label, TypeName(Me)
    OnExecute
End Sub
