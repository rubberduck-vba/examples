VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BindingManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Implements helper methods for the View to consume."
'@Folder MVVM.Infrastructure.Bindings
'@ModuleDescription "Implements helper methods for the View to consume."
'@PredeclaredId
'@Exposed
Option Explicit
Implements IBindingManager
Implements IHandlePropertyChanged

Private Type TState
    Handlers As Collection
    CommandBindings As Collection
    PropertyBindings As Collection
End Type

Private this As TState

Public Function Create() As IBindingManager
    GuardClauses.GuardNonDefaultInstance Me, BindingManager
    Set Create = New BindingManager
End Function

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is BindingManager
End Property

Private Sub ApplyBindings(ByVal Source As Object)
    GuardClauses.GuardExpression IsDefaultInstance, TypeName(Me), "Member call is invalid against stateless default instance."
    Dim Binding As IPropertyBinding
    For Each Binding In this.PropertyBindings
        Binding.Apply
    Next
    OnEvaluateCanExecute Source
End Sub

Private Sub OnEvaluateCanExecute(ByVal Source As Object)
    Dim Binding As ICommandBinding
    For Each Binding In this.CommandBindings
        Binding.EvaluateCanExecute Source
    Next
End Sub

Private Sub Class_Initialize()
    If Not IsDefaultInstance Then
        Set this.Handlers = New Collection
        Set this.CommandBindings = New Collection
        Set this.PropertyBindings = New Collection
    End If
End Sub

Private Sub IBindingManager_BindPropertyPath(ByVal Source As Object, ByVal PropertyPath As String, ByVal Target As Object, _
Optional ByVal TargetProperty As String, _
Optional ByVal Mode As BindingMode, _
Optional ByVal Validator As IValueValidator, _
Optional ByVal Converter As IValueConverter, _
Optional ByVal ErrorFormat As IValidationErrorFormatter)
    
    GuardClauses.GuardExpression IsDefaultInstance, TypeName(Me), "Member call is invalid against stateless default instance."
    
    Dim Binding As IPropertyBinding
    
    Select Case True
    
        Case TypeOf Target Is MSForms.CheckBox And (TargetProperty = "Value" Or TargetProperty = vbNullString)
            Set Binding = PropertyBinding.ForCheckBox(Me, Target, Source, PropertyPath, Converter:=Converter, ErrorFormat:=ErrorFormat)
            
        Case TypeOf Target Is MSForms.ComboBox And (TargetProperty = "Value" Or TargetProperty = vbNullString)
            Set Binding = PropertyBinding.ForComboBoxValue(Me, Target, Source, PropertyPath, Converter:=Converter, ErrorFormat:=ErrorFormat)
            
        Case TypeOf Target Is MSForms.Frame And (TargetProperty = "Caption" Or TargetProperty = vbNullString)
            Set Binding = PropertyBinding.ForFrame(Me, Target, Source, PropertyPath, Converter:=Converter, ErrorFormat:=ErrorFormat)
            
        Case TypeOf Target Is MSForms.Label And (TargetProperty = "Caption" Or TargetProperty = vbNullString)
            Set Binding = PropertyBinding.ForLabel(Me, Target, Source, PropertyPath, Converter:=Converter, ErrorFormat:=ErrorFormat)
            
        Case TypeOf Target Is MSForms.ListBox And TargetProperty = "List"
            Set Binding = PropertyBinding.Create(Me, Target, TargetProperty, Source, PropertyPath, Validator:=Validator, Converter:=Converter)
            
        Case TypeOf Target Is MSForms.ListBox And (TargetProperty = "Value" Or TargetProperty = vbNullString)
            Set Binding = PropertyBinding.ForListBoxValue(Me, Target, Source, PropertyPath, Converter:=Converter)
            
        Case TypeOf Target Is MSForms.OptionButton And (TargetProperty = "Value" Or TargetProperty = vbNullString)
            Set Binding = PropertyBinding.ForOptionButton(Me, Target, Source, PropertyPath, Converter:=Converter, ErrorFormat:=ErrorFormat)
            
        Case TypeOf Target Is MSForms.TextBox And (TargetProperty = "Text" Or TargetProperty = vbNullString)
            Set Binding = PropertyBinding.ForTextBox(Me, Target, Source, PropertyPath, Validator:=Validator, Converter:=Converter, ErrorFormat:=ErrorFormat)
            
        Case Else
            Set Binding = PropertyBinding.Create(Me, Target, TargetProperty, Source, PropertyPath, Mode, OnPropertyChanged, Validator, Converter, ErrorFormat)
            
    End Select
    
    this.PropertyBindings.Add Binding
    this.Handlers.Add Binding
    
    If TargetProperty = vbNullString Then
        Debug.Print TypeName(Source) & ": Binding property path '" & PropertyPath & "' to the default-binding property of " & TypeName(Target) & " object '" & Target.Name & "'."
        
    Else
        Debug.Print TypeName(Source) & ": Binding property path '" & PropertyPath & "' to " & TypeName(Target) & " object '" & Target.Name & "' property '" & TargetProperty & "'."
        
    End If
    
End Sub

Private Sub IBindingManager_BindCommand(ByVal Source As Object, ByVal Target As Object, ByVal Command As ICommand)
    
    GuardClauses.GuardExpression IsDefaultInstance, TypeName(Me), "Member call is invalid against stateless default instance."
    
    Dim Binding As ICommandBinding
    
    Select Case True
    
        Case TypeOf Target Is MSForms.CommandButton
            Set Binding = CommandBinding.ForCommandButton(Target, Command, Source)
            
        Case TypeOf Target Is MSForms.Image
            Set Binding = CommandBinding.ForImage(Target, Command, Source)
            
        Case TypeOf Target Is MSForms.Label
            Set Binding = CommandBinding.ForLabel(Target, Command, Source)
            
        Case TypeOf Target Is MSForms.CheckBox
            Set Binding = CommandBinding.ForCheckBox(Target, Command, Source)
            
        Case Else
            GuardClauses.GuardExpression True, TypeName(Source), "Target type '" & TypeName(Target) & "' does not currently support command bindings."
            
    End Select
    
    this.CommandBindings.Add Binding
    Debug.Print TypeName(Source) & ": Binding command '" & TypeName(Command) & "' to " & TypeName(Target) & " object '" & Target.Name & "'."
    
End Sub

Private Sub IBindingManager_ApplyBindings(ByVal Source As Object)
    GuardClauses.GuardExpression IsDefaultInstance, TypeName(Me), "Member call is invalid against stateless default instance."
    ApplyBindings Source
End Sub

Private Sub IBindingManager_OnEvaluateCanExecute(ByVal Source As Object)
    GuardClauses.GuardExpression IsDefaultInstance, TypeName(Me), "Member call is invalid against stateless default instance."
    OnEvaluateCanExecute Source
End Sub

Private Sub IHandlePropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    OnEvaluateCanExecute Source
End Sub
