VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExampleViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An example ViewModel implementation for some dialog."
'@Folder MVVM.Example
'@ModuleDescription "An example ViewModel implementation for some dialog."
'@PredeclaredId
Implements IViewModel
Implements INotifyPropertyChanged
Option Explicit

Public Event PropertyChanged(ByVal Source As Object, ByVal PropertyName As String)

Private Type TViewModel
    
    'INotifyPropertyChanged state:
    Handlers As Collection
    
    'CommandBindings:
    SomeCommand As ICommand
    
    'Read/Write PropertyBindings:
    SourcePath As String
    SomeOption As Boolean
    SomeOtherOption As Boolean
    
End Type

Private this As TViewModel
Private WithEvents ValidationHandler As ValidationManager
Attribute ValidationHandler.VB_VarHelpID = -1

Public Function Create() As IViewModel
    GuardClauses.GuardNonDefaultInstance Me, ExampleViewModel, TypeName(Me)
    
    Dim result As ExampleViewModel
    Set result = New ExampleViewModel
    
    Set Create = result
End Function

Public Property Get Validation() As IHandleValidationError
    Set Validation = ValidationHandler
End Property

Public Property Get SourcePath() As String
    SourcePath = this.SourcePath
End Property

Public Property Let SourcePath(ByVal RHS As String)
    If this.SourcePath <> RHS Then
        this.SourcePath = RHS
        OnPropertyChanged "SourcePath"
    End If
End Property

Public Property Get SomeOption() As Boolean
    SomeOption = this.SomeOption
End Property

Public Property Let SomeOption(ByVal RHS As Boolean)
    If this.SomeOption <> RHS Then
        this.SomeOption = RHS
        OnPropertyChanged "SomeOption"
    End If
End Property

Public Property Get SomeOtherOption() As Boolean
    SomeOtherOption = this.SomeOtherOption
End Property

Public Property Let SomeOtherOption(ByVal RHS As Boolean)
    If this.SomeOtherOption <> RHS Then
        this.SomeOtherOption = RHS
        OnPropertyChanged "SomeOtherOption"
    End If
End Property

Public Property Get SomeCommand() As ICommand
    Set SomeCommand = this.SomeCommand
End Property

Public Property Set SomeCommand(ByVal RHS As ICommand)
    Set this.SomeCommand = RHS
End Property

Public Property Get SomeOptionName() As String
    SomeOptionName = "Auto"
End Property

Public Property Get SomeOtherOptionName() As String
    SomeOtherOptionName = "Manual/Browse"
End Property

Public Property Get Instructions() As String
    Instructions = "Lorem ipsum dolor sit amet, consectetur adipiscing elit."
End Property

Private Sub OnPropertyChanged(ByVal PropertyName As String)
    RaiseEvent PropertyChanged(Me, PropertyName)
    Dim Handler As IHandlePropertyChanged
    For Each Handler In this.Handlers
        Handler.OnPropertyChanged Me, PropertyName
    Next
End Sub

Private Sub Class_Initialize()
    Set this.Handlers = New Collection
    Set ValidationHandler = ValidationManager.Create
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    OnPropertyChanged PropertyName
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    this.Handlers.Add Handler
End Sub

Private Property Get IViewModel_Validation() As IHandleValidationError
    Set IViewModel_Validation = ValidationHandler
End Property

Private Sub ValidationHandler_PropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    OnPropertyChanged PropertyName
End Sub
