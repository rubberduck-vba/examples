Attribute VB_Name = "PropertyBindings"
Option Explicit

'@Description "Binds a MSForms.Control property to a source property"
Public Function BindProperty(ByVal Control As MSForms.Control, ByVal ControlProperty As String, ByVal SourceProperty As String, ByVal Source As Object, Optional ByVal InvertBoolean As Boolean = False) As OneWayPropertyBinding
    
    Dim Binding As OneWayPropertyBinding
    Set Binding = New OneWayPropertyBinding
    
    Binding.Initialize Control, ControlProperty, Source, SourceProperty, InvertBoolean
    
    Set BindProperty = Binding

End Function

'@Description "Binds the Text/Value of a MSForms.TextBox to a source property"
Public Function BindTextBox(ByVal Control As MSForms.TextBox, ByVal SourceProperty As String, ByVal Source As Object) As TextBoxValueBinding
    
    Dim Binding As TextBoxValueBinding
    Set Binding = New TextBoxValueBinding
    
    Binding.Initialize Control, Source, SourceProperty
    
    Set BindTextBox = Binding
    
End Function

'@Description "Binds the Text of a MSForms.ComboBox to a String source property"
Public Function BindComboBox(ByVal Control As MSForms.ComboBox, ByVal SourceProperty As String, ByVal Source As Object) As ComboBoxValueBinding
    
    Dim Binding As ComboBoxValueBinding
    Set Binding = New ComboBoxValueBinding
    
    Binding.Initialize Control, Source, SourceProperty
    
    Set BindComboBox = Binding

End Function

'@Description "Binds the Value of a MSForms.CheckBox to a Boolean source property"
Public Function BindCheckBox(ByVal Control As MSForms.CheckBox, ByVal SourceProperty As String, ByVal Source As Object) As CheckBoxValueBinding
    
    Dim Binding As CheckBoxValueBinding
    Set Binding = New CheckBoxValueBinding
    
    Binding.Initialize Control, Source, SourceProperty
    
    Set BindCheckBox = Binding

End Function

