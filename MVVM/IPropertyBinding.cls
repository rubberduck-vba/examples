VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for binding a ViewModel property path to a UI element."
'@Folder MVVM.Infrastructure.Bindings
'@ModuleDescription "An object responsible for binding a ViewModel property path to a UI element."
'@Interface
'@Exposed
Option Explicit

Public Enum BindingMode
    TwoWayBinding
    OneWayBinding
    OneWayToSource
    OneTimeBinding
End Enum

Public Enum BindingUpdateSourceTrigger
    OnPropertyChanged
    OnKeyPress
    OnExit
End Enum

'@Description "Gets a value indicating the binding mode/direction."
Public Property Get Mode() As BindingMode
Attribute Mode.VB_Description = "Gets a value indicating the binding mode/direction."
End Property

'@Description "Gets a value indicating the binding update trigger."
Public Property Get UpdateSourceTrigger() As BindingUpdateSourceTrigger
Attribute UpdateSourceTrigger.VB_Description = "Gets a value indicating the binding update trigger."
End Property

'@Description "Gets the ViewModel object that is the binding source."
Public Property Get Source() As Object
Attribute Source.VB_Description = "Gets the ViewModel object that is the binding source."
End Property

'@Description "Gets a path to a Public Property Get procedure resolvable from the binding source."
Public Property Get SourcePropertyPath() As String
Attribute SourcePropertyPath.VB_Description = "Gets a path to a Public Property Get procedure resolvable from the binding source."
End Property

'@Description "Gets the UI element that is the binding target."
Public Property Get Target() As Object
Attribute Target.VB_Description = "Gets the UI element that is the binding target."
End Property

'@Description "Gets the name of the bound property on the target object."
Public Property Get TargetProperty() As String
Attribute TargetProperty.VB_Description = "Gets the name of the bound property on the target object."
End Property

'@Description "Gets the converter to use (if any) when applying the binding."
Public Property Get Converter() As IValueConverter
Attribute Converter.VB_Description = "Gets the converter to use (if any) when applying the binding."
End Property

'@Description "Applies the binding."
Public Sub Apply()
Attribute Apply.VB_Description = "Applies the binding."
End Sub

