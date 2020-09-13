VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBindingManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Exposes helper methods for the View to consume."
'@Folder MVVM.Infrastructure.Bindings
'@ModuleDescription "Exposes helper methods for the View to consume."
'@Interface
'@Exposed
Option Explicit

'@Description "Creates a new property binding for the specified path, source, and target."
Public Sub BindPropertyPath(ByVal Source As Object, ByVal PropertyPath As String, ByVal Target As Object, Optional ByVal TargetProperty As String, Optional ByVal Mode As BindingMode, Optional ByVal Validator As IValueValidator, Optional ByVal Converter As IValueConverter, Optional ByVal ErrorFormat As IValidationErrorFormatter)
Attribute BindPropertyPath.VB_Description = "Creates a new property binding for the specified path, source, and target."
End Sub

'@Description "Binds the specified command to user interaction with the specified target UI element (e.g. a MSForms.CommandButton control)."
Public Sub BindCommand(ByVal Source As Object, ByVal Target As Object, ByVal Command As ICommand)
Attribute BindCommand.VB_Description = "Binds the specified command to user interaction with the specified target UI element (e.g. a MSForms.CommandButton control)."
End Sub

'@Description "Applies all property bindings (side effect: commands re-evaluate whether they can execute)."
Public Sub ApplyBindings(ByVal Source As Object)
Attribute ApplyBindings.VB_Description = "Applies all property bindings (side effect: commands re-evaluate whether they can execute)."
End Sub

'@Description "Causes all commands to evaluate whether they can be executed."
Public Sub OnEvaluateCanExecute(ByVal Source As Object)
Attribute OnEvaluateCanExecute.VB_Description = "Causes all commands to evaluate whether they can be executed."
End Sub
