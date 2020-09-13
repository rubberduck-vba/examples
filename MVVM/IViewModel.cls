VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Formalizes the basic functionality of a ViewModel."
'@Exposed
'@Folder MVVM.Infrastructure
'@ModuleDescription "Formalizes the basic functionality of a ViewModel."
'@Interface
Option Explicit

'@Description "Gets the validation error handler for this ViewModel."
Public Property Get Validation() As IHandleValidationError
Attribute Validation.VB_Description = "Gets the validation error handler for this ViewModel."
End Property
