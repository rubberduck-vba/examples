VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHandleValidationError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object that can be registered as a listener for data validation errors."
'@Folder MVVM.Infrastructure.Bindings.Validation
'@ModuleDescription "An object that can be registered as a listener for data validation errors."
'@Interface
'@Exposed
Option Explicit

'@Description "Resets the IsValid state."
Public Sub ClearValidationErrors(ByVal PropertyName As String)
Attribute ClearValidationErrors.VB_Description = "Resets the IsValid state."
End Sub

'@Description "Gets an array containing all validation error messages."
Public Property Get ValidationErrors() As Variant
Attribute ValidationErrors.VB_Description = "Gets an array containing all validation error messages."
End Property

'@Description "A callback invoked when a validation error occurs as a property binding is updated."
Public Sub OnValidationError(ByVal Source As IPropertyBinding, ByVal Message As String)
Attribute OnValidationError.VB_Description = "A callback invoked when a validation error occurs as a property binding is updated."
End Sub

'@Description "Gets a flag that is True unless the implementing ViewModel has validation errors."
Public Property Get IsValid() As Boolean
Attribute IsValid.VB_Description = "Gets a flag that is True unless the implementing ViewModel has validation errors."
End Property

'@Description "Gets the first validation error message."
Public Property Get Message() As String
Attribute Message.VB_Description = "Gets the first validation error message."
End Property
