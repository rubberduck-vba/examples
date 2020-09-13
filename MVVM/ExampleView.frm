VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExampleView 
   Caption         =   "ExampleView"
   ClientHeight    =   3084
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   3708
   OleObjectBlob   =   "ExampleView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExampleView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An example implementation of a View."

'@Folder MVVM.Example
'@ModuleDescription "An example implementation of a View."
Implements IView
Implements ICancellable
Option Explicit

Private Type TView
    'IView state:
    ViewModel As ExampleViewModel
    
    'ICancellable state:
    IsCancelled As Boolean
    
    'Data binding helper dependency:
    Bindings As IBindingManager
End Type

Private this As TView

'@Description "A factory method to create new instances of this View, already wired-up to a ViewModel."
Public Function Create(ByVal ViewModel As ExampleViewModel, ByVal Bindings As IBindingManager) As IView
Attribute Create.VB_Description = "A factory method to create new instances of this View, already wired-up to a ViewModel."
    GuardClauses.GuardNonDefaultInstance Me, ExampleView, TypeName(Me)
    GuardClauses.GuardNullReference ViewModel, TypeName(Me)
    GuardClauses.GuardNullReference Bindings, TypeName(Me)
    
    Dim result As ExampleView
    Set result = New ExampleView
    
    Set result.Bindings = Bindings
    Set result.ViewModel = ViewModel
    
    Set Create = result
    
End Function

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is ExampleView
End Property

'@Description "Gets/sets the ViewModel to use as a context for property and command bindings."
Public Property Get ViewModel() As ExampleViewModel
Attribute ViewModel.VB_Description = "Gets/sets the ViewModel to use as a context for property and command bindings."
    Set ViewModel = this.ViewModel
End Property

Public Property Set ViewModel(ByVal RHS As ExampleViewModel)
    GuardClauses.GuardExpression IsDefaultInstance, TypeName(Me)
    GuardClauses.GuardNullReference RHS
    
    Set this.ViewModel = RHS
    InitializeBindings

End Property

'@Description "Gets/sets the binding manager implementation."
Public Property Get Bindings() As IBindingManager
Attribute Bindings.VB_Description = "Gets/sets the binding manager implementation."
    Set Bindings = this.Bindings
End Property

Public Property Set Bindings(ByVal RHS As IBindingManager)
    GuardClauses.GuardExpression IsDefaultInstance, TypeName(Me)
    GuardClauses.GuardDoubleInitialization this.Bindings, TypeName(Me)
    GuardClauses.GuardNullReference RHS
    
    Set this.Bindings = RHS

End Property

Private Sub BindViewModelCommands()
    With Bindings
        .BindCommand ViewModel, Me.OkButton, AcceptCommand.Create(Me)
        .BindCommand ViewModel, Me.CancelButton, CancelCommand.Create(Me)
        .BindCommand ViewModel, Me.BrowseButton, ViewModel.SomeCommand
        '...
    End With
End Sub

Private Sub BindViewModelProperties()
    With Bindings
        
        .BindPropertyPath ViewModel, "SourcePath", Me.PathBox, _
            Validator:=New RequiredStringValidator, _
            ErrorFormat:=AggregateErrorFormatter.Create(ViewModel, _
                ValidationErrorFormatter.Create(Me.PathBox).WithErrorBackgroundColor.WithErrorBorderColor, _
                ValidationErrorFormatter.Create(Me.InvalidPathIcon).WithTargetOnlyVisibleOnError("SourcePath"), _
                ValidationErrorFormatter.Create(Me.ValidationMessage1).WithTargetOnlyVisibleOnError("SourcePath"))
        
        .BindPropertyPath ViewModel, "Instructions", Me.InstructionsLabel
        
        .BindPropertyPath ViewModel, "SomeOption", Me.OptionButton1
        .BindPropertyPath ViewModel, "SomeOtherOption", Me.OptionButton2
        .BindPropertyPath ViewModel, "SomeOptionName", Me.OptionButton1, "Caption", OneTimeBinding
        .BindPropertyPath ViewModel, "SomeOtherOptionName", Me.OptionButton2, "Caption", OneTimeBinding
        
        '...
        
    End With
End Sub

Private Sub InitializeBindings()
    If ViewModel Is Nothing Then Exit Sub
    BindViewModelProperties
    BindViewModelCommands
    Bindings.ApplyBindings ViewModel
End Sub

Private Sub OnCancel()
    this.IsCancelled = True
    Me.Hide
End Sub

Private Property Get ICancellable_IsCancelled() As Boolean
    ICancellable_IsCancelled = this.IsCancelled
End Property

Private Sub ICancellable_OnCancel()
    OnCancel
End Sub

Private Sub IView_Hide()
    Me.Hide
End Sub

Private Sub IView_Show()
    Me.Show vbModal
End Sub

Private Function IView_ShowDialog() As Boolean
    Me.Show vbModal
    IView_ShowDialog = Not this.IsCancelled
End Function

Private Property Get IView_ViewModel() As Object
    Set IView_ViewModel = this.ViewModel
End Property
