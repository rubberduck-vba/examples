VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for binding a ViewModel property path to a UI element."
'@Folder MVVM.Infrastructure.Bindings
'@ModuleDescription "An object responsible for binding a ViewModel property path to a UI element."
'@PredeclaredId
'@Exposed
Implements IPropertyBinding
Implements IHandlePropertyChanged
Implements INotifyPropertyChanged
Option Explicit

Private Type TState
    Handlers As Collection
    Manager As IBindingManager
    Mode As BindingMode
    UpdateSourceTrigger As BindingUpdateSourceTrigger
    Source As Object
    SourcePropertyPath As String
    Target As Object
    TargetProperty As String
    Converter As IValueConverter
    Validator As IValueValidator
    ValidationErrorHandler As IHandleValidationError
    ValidationErrorFormatter As IValidationErrorFormatter
    Applied As Boolean
End Type

Private WithEvents ControlEventSource As MSForms.Control
Attribute ControlEventSource.VB_VarHelpID = -1
Private WithEvents TextBoxEventSource As MSForms.TextBox
Attribute TextBoxEventSource.VB_VarHelpID = -1
Private WithEvents CheckBoxEventSource As MSForms.CheckBox
Attribute CheckBoxEventSource.VB_VarHelpID = -1
Private WithEvents OptionButtonEventSource As MSForms.OptionButton
Attribute OptionButtonEventSource.VB_VarHelpID = -1
Private WithEvents ComboBoxEventSource As MSForms.ComboBox
Attribute ComboBoxEventSource.VB_VarHelpID = -1
Private WithEvents ListBoxEventSource As MSForms.ListBox
Attribute ListBoxEventSource.VB_VarHelpID = -1

Private this As TState

Public Function ForTextBox(ByVal Manager As IBindingManager, ByVal Target As MSForms.TextBox, ByVal Source As Object, ByVal SourceProperty As String, _
Optional ByVal Mode As BindingMode = BindingMode.TwoWayBinding, _
Optional ByVal UpdateSource As BindingUpdateSourceTrigger = OnPropertyChanged, _
Optional ByVal Validator As IValueValidator, _
Optional ByVal Converter As IValueConverter, _
Optional ByVal ErrorFormat As IValidationErrorFormatter) As IPropertyBinding
    Set ForTextBox = Create(Manager, Target, "Text", Source, SourceProperty, Mode, UpdateSource, Validator, Converter, ErrorFormat)
End Function

Public Function ForCheckBox(ByVal Manager As IBindingManager, ByVal Target As MSForms.CheckBox, ByVal Source As Object, ByVal SourceProperty As String, _
Optional ByVal Mode As BindingMode = BindingMode.TwoWayBinding, _
Optional ByVal Validator As IValueValidator, _
Optional ByVal Converter As IValueConverter, _
Optional ByVal ErrorFormat As IValidationErrorFormatter) As IPropertyBinding
    Set ForCheckBox = Create(Manager, Target, "Value", Source, SourceProperty, Mode, OnPropertyChanged, Validator, Converter, ErrorFormat)
End Function

Public Function ForOptionButton(ByVal Manager As IBindingManager, ByVal Target As MSForms.OptionButton, ByVal Source As Object, ByVal SourceProperty As String, _
Optional ByVal Mode As BindingMode = BindingMode.TwoWayBinding, _
Optional ByVal Validator As IValueValidator, _
Optional ByVal Converter As IValueConverter, _
Optional ByVal ErrorFormat As IValidationErrorFormatter) As IPropertyBinding
    Set ForOptionButton = Create(Manager, Target, "Value", Source, SourceProperty, Mode, OnPropertyChanged, Validator, Converter, ErrorFormat)
End Function

Public Function ForLabel(ByVal Manager As IBindingManager, ByVal Target As MSForms.Label, ByVal Source As Object, ByVal SourceProperty As String, _
Optional ByVal Mode As BindingMode = BindingMode.OneWayBinding, _
Optional ByVal Validator As IValueValidator, _
Optional ByVal Converter As IValueConverter, _
Optional ByVal ErrorFormat As IValidationErrorFormatter) As IPropertyBinding
    Set ForLabel = Create(Manager, Target, "Caption", Source, SourceProperty, Mode, OnPropertyChanged, Validator, Converter, ErrorFormat)
End Function

Public Function ForFrame(ByVal Manager As IBindingManager, ByVal Target As MSForms.Frame, ByVal Source As Object, ByVal SourceProperty As String, _
Optional ByVal Mode As BindingMode = BindingMode.OneWayBinding, _
Optional ByVal Validator As IValueValidator, _
Optional ByVal Converter As IValueConverter, _
Optional ByVal ErrorFormat As IValidationErrorFormatter) As IPropertyBinding
    Set ForFrame = Create(Manager, Target, "Caption", Source, SourceProperty, Mode, OnPropertyChanged, Validator, Converter, ErrorFormat)
End Function

Public Function ForComboBoxValue(ByVal Manager As IBindingManager, ByVal Target As MSForms.ComboBox, ByVal Source As Object, ByVal SourceProperty As String, _
Optional ByVal Mode As BindingMode = BindingMode.TwoWayBinding, _
Optional ByVal Validator As IValueValidator, _
Optional ByVal Converter As IValueConverter, _
Optional ByVal ErrorFormat As IValidationErrorFormatter) As IPropertyBinding
    Set ForComboBoxValue = Create(Manager, Target, "Value", Source, SourceProperty, Mode, OnPropertyChanged, Validator, Converter, ErrorFormat)
End Function

Public Function ForListBoxValue(ByVal Manager As IBindingManager, ByVal Target As MSForms.ListBox, ByVal Source As Object, ByVal SourceProperty As String, Optional ByVal Mode As BindingMode = BindingMode.TwoWayBinding, Optional ByVal Validator As IValueValidator, Optional ByVal Converter As IValueConverter) As IPropertyBinding
    Set ForListBoxValue = Create(Manager, Target, "Value", Source, SourceProperty, Mode, OnPropertyChanged, Validator, Converter)
End Function

Public Function Create(ByVal Manager As IBindingManager, ByVal Target As Object, ByVal TargetProperty As String, ByVal Source As Object, ByVal SourceProperty As String, _
Optional ByVal Mode As BindingMode = BindingMode.TwoWayBinding, _
Optional ByVal UpdateSource As BindingUpdateSourceTrigger = BindingUpdateSourceTrigger.OnPropertyChanged, _
Optional ByVal Validator As IValueValidator, _
Optional ByVal Converter As IValueConverter, _
Optional ByVal ErrorFormat As IValidationErrorFormatter) As IPropertyBinding
    
    Dim result As PropertyBinding
    Set result = New PropertyBinding
    
    Set result.Target = Target
    result.TargetProperty = TargetProperty
    
    Set result.Source = Source
    result.SourcePropertyPath = SourceProperty
    
    Set result.Manager = Manager
    Set result.Converter = Converter
    Set result.Validator = Validator
    Set result.ValidationErrorFormatter = ErrorFormat
    result.Mode = Mode
    result.UpdateSourceTrigger = UpdateSource
    
    If Mode <> OneWayToSource And TypeOf Source Is INotifyPropertyChanged Then
        Dim notifier As INotifyPropertyChanged
        Set notifier = Source
        notifier.RegisterHandler result
    End If
    
    Set Create = result
    
End Function

Public Property Get Source() As Object
    Set Source = this.Source
End Property

Public Property Set Source(ByVal RHS As Object)
    GuardClauses.GuardDoubleInitialization this.Manager, TypeName(Me)
    GuardClauses.GuardNullReference RHS, TypeName(Me)
    Set this.Source = RHS
    If TypeOf RHS Is IViewModel Then
        Dim ViewModel As IViewModel
        Set ViewModel = RHS
        Set this.ValidationErrorHandler = ViewModel.Validation
    End If
End Property

Public Property Get Target() As Object
    Set Target = this.Target
End Property

Public Property Set Target(ByVal RHS As Object)
    GuardClauses.GuardDoubleInitialization this.Target, TypeName(Me)
    GuardClauses.GuardNullReference RHS, TypeName(Me)
    
    Set this.Target = RHS
    
    Set ControlEventSource = Nothing
    Set TextBoxEventSource = Nothing
    Set CheckBoxEventSource = Nothing
    Set ComboBoxEventSource = Nothing
    Set ListBoxEventSource = Nothing
    Set OptionButtonEventSource = Nothing
    
    If TypeOf this.Target Is MSForms.Control And this.UpdateSourceTrigger = OnExit Then Set ControlEventSource = this.Target
    
    Select Case True
    
        Case TypeOf this.Target Is MSForms.TextBox
            Set TextBoxEventSource = this.Target
            
        Case TypeOf this.Target Is MSForms.CheckBox And TypeName(this.Target) = "CheckBox" 'OptionButton matches the interface criteria
            Set CheckBoxEventSource = this.Target
            
        Case TypeOf this.Target Is MSForms.ComboBox
            Set ComboBoxEventSource = this.Target
            
        Case TypeOf this.Target Is MSForms.ListBox
            Set ListBoxEventSource = this.Target
            
        Case TypeOf this.Target Is MSForms.OptionButton
            Set OptionButtonEventSource = this.Target
            
    End Select
    
End Property

Public Property Get SourcePropertyPath() As String
    SourcePropertyPath = this.SourcePropertyPath
End Property

Public Property Let SourcePropertyPath(ByVal RHS As String)
    this.SourcePropertyPath = RHS
End Property

Public Property Get TargetProperty() As String
    TargetProperty = this.TargetProperty
End Property

Public Property Let TargetProperty(ByVal RHS As String)
    this.TargetProperty = RHS
End Property

Public Property Get Mode() As BindingMode
    Mode = this.Mode
End Property

Public Property Let Mode(ByVal RHS As BindingMode)
    this.Mode = RHS
End Property

Public Property Get UpdateSourceTrigger() As BindingUpdateSourceTrigger
    UpdateSourceTrigger = this.UpdateSourceTrigger
End Property

Public Property Let UpdateSourceTrigger(ByVal RHS As BindingUpdateSourceTrigger)
    this.UpdateSourceTrigger = RHS
End Property

Public Property Get Validator() As IValueValidator
    Set Validator = this.Validator
End Property

Public Property Set Validator(ByVal RHS As IValueValidator)
    Set this.Validator = RHS
End Property

Public Property Get ValidationErrorFormatter() As IValidationErrorFormatter
    Set ValidationErrorFormatter = this.ValidationErrorFormatter
End Property

Public Property Set ValidationErrorFormatter(ByVal RHS As IValidationErrorFormatter)
    Set this.ValidationErrorFormatter = RHS
End Property

Public Property Get Converter() As IValueConverter
    Set Converter = this.Converter
End Property

Public Property Set Converter(ByVal RHS As IValueConverter)
    Set this.Converter = RHS
End Property

Public Property Get Manager() As IBindingManager
    Set Manager = this.Manager
End Property

Public Property Set Manager(ByVal RHS As IBindingManager)
    GuardClauses.GuardDoubleInitialization this.Manager, TypeName(Me)
    GuardClauses.GuardNullReference RHS, TypeName(Me)
    Set this.Manager = RHS
End Property

Private Function TryReadPropertyValue(ByVal Instance As Object, ByVal PropertyName As String, ByRef outValue As Variant) As Boolean
    On Error Resume Next
    outValue = CallByName(Instance, PropertyName, VbGet)
    TryReadPropertyValue = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function TryWritePropertyValue(ByVal Instance As Object, ByVal PropertyName As String, ByVal Value As Variant) As Boolean
    On Error Resume Next
    CallByName Instance, PropertyName, VbLet, Value
    TryWritePropertyValue = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Sub ApplyToSource()
'reads from the target and writes to the source.

    'one-time and one-way binding modes do not write values to the binding source (ViewModel).
    If this.Mode = OneTimeBinding Or this.Mode = OneWayBinding Then Exit Sub
    
    Dim Target As Object
    Set Target = ResolvePropertyPath(this.Target, this.TargetProperty)
    
    Dim TargetPropertyName As String
    TargetPropertyName = ResolvePropertyName(this.TargetProperty)
    
    Dim TargetValue As Variant
    If Not TryReadPropertyValue(Target, TargetPropertyName, outValue:=TargetValue) Then
        Debug.Print TypeName(Me) & ":[" & ToString & "] failed to read target property value."
        Exit Sub
    End If
    
    If Not this.Converter Is Nothing Then TargetValue = this.Converter.ConvertBack(TargetValue)
    Validate TargetValue
    
    Dim Source As Object
    Set Source = ResolvePropertyPath(this.Source, this.SourcePropertyPath)
    
    Dim SourcePropertyName As String
    SourcePropertyName = ResolvePropertyName(this.SourcePropertyPath)
    
    Dim CurrentValue As Variant
    If Not TryReadPropertyValue(Source, SourcePropertyName, outValue:=CurrentValue) Then
        Debug.Print TypeName(Me) & ":[" & ToString & "] failed to read current source property value."
        Exit Sub
    End If
    
    If TargetValue <> CurrentValue Then
        
        If Not this.ValidationErrorHandler Is Nothing Then this.ValidationErrorHandler.ClearValidationErrors SourcePropertyName
        If Not TryWritePropertyValue(Source, SourcePropertyName, TargetValue) Then
            Debug.Print TypeName(Me) & ":[" & ToString & "] failed to write to source property."
            Exit Sub
        End If
        
        Debug.Print TypeName(Me) & ":[" & ToString & "] was successfully applied to source."
        this.Manager.OnEvaluateCanExecute this.Source
        
    End If
    
End Sub

Private Function ResolvePropertyPath(ByVal Source As Object, ByVal PropertyPath As String) As Object
    
    Dim Parts As Variant
    Parts = Strings.Split(PropertyPath, ".")
    
    If UBound(Parts) = LBound(Parts) Then
        Set ResolvePropertyPath = Source
    Else
        Dim RecursiveProperty As Object
        Set RecursiveProperty = CallByName(Source, Parts(0), VbGet)
        If RecursiveProperty Is Nothing Then Exit Function
        Set ResolvePropertyPath = ResolvePropertyPath(RecursiveProperty, Right$(PropertyPath, Len(PropertyPath) - Len(Parts(0)) - 1))
    End If
    
End Function

Private Function ResolvePropertyName(ByVal PropertyPath As String) As String
    Dim Parts As Variant
    Parts = Strings.Split(PropertyPath, ".")
    ResolvePropertyName = Parts(UBound(Parts))
End Function

Private Sub Apply()
'reads from the source and writes to the target.
    
    'one-way to source and already-applied one-time binding modes do not apply to target
    If this.Mode = OneWayToSource Or (this.Mode = OneTimeBinding And this.Applied) Then Exit Sub
    
    Dim Source As Object
    Set Source = ResolvePropertyPath(this.Source, this.SourcePropertyPath)
    
    Dim SourceValue As Variant
    Dim UseFallbackDefault As Boolean
    
    If Source Is Nothing Then
        UseFallbackDefault = TryGetDefaultBindingValue(outValue:=SourceValue)
    End If
    
    If Source Is Nothing And Not UseFallbackDefault Then
        Debug.Print "Cannot bind target property '" & this.TargetProperty & "' (" & TypeName(this.Target) & ": '" & this.Target.Name & "'); source object in path '" & this.SourcePropertyPath & "' is Nothing"
        Exit Sub
    ElseIf UseFallbackDefault Then
        Debug.Print "Source object in path '" & this.SourcePropertyPath & "' is Nothing; binding target property '" & this.TargetProperty & "' to default/fallback value."
    End If
    
    Dim SourcePropertyName As String
    SourcePropertyName = ResolvePropertyName(this.SourcePropertyPath)
    
    On Error Resume Next
    If Not UseFallbackDefault Then
        SourceValue = CallByName(Source, SourcePropertyName, VbGet)
        If Err.Number <> 0 Then
            Debug.Print "Failed to apply binding for target property '" & this.TargetProperty & "' (" & TypeName(this.Target) & ": '" & this.Target.Name & "'); source path: '" & this.SourcePropertyPath & "'; " & Err.Description
            On Error GoTo 0
            Err.Clear
            Exit Sub
        End If
    End If
    On Error GoTo 0
    
    Validate SourceValue
    
    Dim Target As Object
    Set Target = ResolvePropertyPath(this.Target, this.TargetProperty)
    Debug.Assert Not Target Is Nothing
    
    If Not this.Converter Is Nothing Then
        SourceValue = this.Converter.Convert(SourceValue)
        'TODO: handle converter errors?
    End If
    
    Dim TargetPropertyName As String
    TargetPropertyName = ResolvePropertyName(this.TargetProperty)
    
    Dim CurrentValue As Variant
    CurrentValue = CallByName(Target, TargetPropertyName, VbGet)
    
    If SourceValue <> CurrentValue Then
        On Error Resume Next
        CallByName Target, TargetPropertyName, VbLet, SourceValue
        If Err.Number <> 0 Then
            Debug.Print "Failed to apply binding for target property '" & this.TargetProperty & "' (source path: " & this.SourcePropertyPath & "): " & Err.Description
            Err.Clear
        Else
            Debug.Print "Successfully updated binding for target property '" & this.TargetProperty & "' (source path: " & this.SourcePropertyPath & ")"
            If Not this.ValidationErrorHandler Is Nothing Then this.ValidationErrorHandler.ClearValidationErrors SourcePropertyName
            this.Manager.OnEvaluateCanExecute this.Source
        End If
        On Error GoTo 0
    End If
End Sub

Private Function TryGetDefaultBindingValue(ByRef outValue As Variant) As Boolean
'Gets a default value for certain specific target properties, used when source path cannot be fully resolved,
'e.g. when target binds to "SomeObjectProperty.SomeProperty" and "SomeObjectProperty" is Nothing.

    Select Case ResolvePropertyName(this.TargetProperty)
    
        Case "Text", "Caption"
            outValue = vbNullString
            TryGetDefaultBindingValue = True
            
        Case "Enabled", "Visible"
            outValue = False
            TryGetDefaultBindingValue = True
            
        Case "Value"
            If TypeOf this.Target Is MSForms.CheckBox _
                Or TypeOf this.Target Is MSForms.OptionButton _
            Then
                outValue = False
                TryGetDefaultBindingValue = True
            End If
            
    End Select
End Function

Private Sub Validate(ByVal TargetValue As Variant)
    If this.Validator Is Nothing Then Exit Sub
    
    If this.Validator.IsValid(TargetValue, this.Source, this.Target) Then
        Debug.Print TypeName(Me) & ":[" & ToString & "] value passed validation."
        If Not this.ValidationErrorFormatter Is Nothing Then
            this.ValidationErrorFormatter.Restore
        End If
        Exit Sub
    
    End If
    
    Debug.Print TypeName(Me) & ":[" & ToString & "] value failed validation and will not be applied. A validation error will be propagated."
    If Not this.ValidationErrorFormatter Is Nothing Then
        this.ValidationErrorFormatter.Apply ResolvePropertyName(this.SourcePropertyPath), ValidationErrorMessage
    End If

    NotifyValidationError
    'Err.Raise ValidationError, TypeName(Me), ValidationErrorMessage
End Sub

Private Property Get ValidationErrorMessage() As String
    If Not this.Validator Is Nothing Then
        ValidationErrorMessage = this.Validator.Message
    'Else: no validator -> value is always valid -> no need for a validation error message
    End If
End Property

Private Sub NotifyValidationError()
    If this.ValidationErrorHandler Is Nothing Then Exit Sub
    this.ValidationErrorHandler.OnValidationError Me, ValidationErrorMessage
End Sub

Private Sub NotifyPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    Dim Handler As IHandlePropertyChanged
    For Each Handler In this.Handlers
        Handler.OnPropertyChanged Source, PropertyName
    Next
End Sub

Private Sub Class_Initialize()
    If Not Me Is PropertyBinding Then Set this.Handlers = New Collection
End Sub

Private Sub IHandlePropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    If PropertyName = ResolvePropertyName(this.SourcePropertyPath) Then Apply
    NotifyPropertyChanged Source, PropertyName
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    NotifyPropertyChanged Source, PropertyName
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    this.Handlers.Add Handler
End Sub

Private Sub IPropertyBinding_Apply()
    On Error GoTo ValidationFailed
    Apply
    Exit Sub
ValidationFailed:
    Debug.Print TypeName(Me) & ": validation failed, binding was not applied."
End Sub

Private Property Get IPropertyBinding_Converter() As IValueConverter
    Set IPropertyBinding_Converter = this.Converter
End Property

Private Property Get IPropertyBinding_Source() As Object
    Set IPropertyBinding_Source = this.Source
End Property

Private Property Get IPropertyBinding_SourcePropertyPath() As String
    IPropertyBinding_SourcePropertyPath = this.SourcePropertyPath
End Property

Private Property Get IPropertyBinding_Target() As Object
    Set IPropertyBinding_Target = this.Target
End Property

Private Property Get IPropertyBinding_TargetProperty() As String
    IPropertyBinding_TargetProperty = this.TargetProperty
End Property

Private Property Get IPropertyBinding_Mode() As BindingMode
    IPropertyBinding_Mode = this.Mode
End Property

Private Property Get IPropertyBinding_UpdateSourceTrigger() As BindingUpdateSourceTrigger
    IPropertyBinding_UpdateSourceTrigger = this.UpdateSourceTrigger
End Property

Private Sub CheckBoxEventSource_Change()
    On Error GoTo FailedValidation
    If this.UpdateSourceTrigger = OnPropertyChanged Then ApplyToSource
    Exit Sub
FailedValidation:
    On Error GoTo 0
End Sub

Private Sub ComboBoxEventSource_Change()
    On Error GoTo FailedValidation
    If this.UpdateSourceTrigger = OnPropertyChanged Then ApplyToSource
    Exit Sub
FailedValidation:
    On Error GoTo 0
End Sub

Private Sub ListBoxEventSource_Change()
    On Error GoTo FailedValidation
    If this.UpdateSourceTrigger = OnPropertyChanged Then ApplyToSource
    Exit Sub
FailedValidation:
    On Error GoTo 0
End Sub

Private Sub OptionButtonEventSource_Change()
    On Error GoTo FailedValidation
    If this.UpdateSourceTrigger = OnPropertyChanged Then ApplyToSource
    Exit Sub
FailedValidation:
    On Error GoTo 0
End Sub

Private Sub TextBoxEventSource_Change()
    On Error GoTo FailedValidation
    If this.UpdateSourceTrigger = OnPropertyChanged Then ApplyToSource
    Exit Sub
FailedValidation:
    On Error GoTo 0
End Sub

Private Sub TextBoxEventSource_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo FailedValidation
    If this.UpdateSourceTrigger = OnKeyPress Then ApplyToSource
    Exit Sub
FailedValidation:
    KeyCode.Value = 0 'swallow invalid keypress
    On Error GoTo 0
End Sub

Private Sub ControlEventSource_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo FailedValidation
    If this.UpdateSourceTrigger = OnExit Then ApplyToSource
    Exit Sub
FailedValidation:
    Cancel.Value = True 'keep target control focused
    On Error GoTo 0
End Sub

Private Function ToString() As String
    ToString = "source path '" & this.SourcePropertyPath & "' (" & TypeName(this.Source) & "), target property '" & this.TargetProperty & "' (" & TypeName(this.Target) & ")"
End Function

