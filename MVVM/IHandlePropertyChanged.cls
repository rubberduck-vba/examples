VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHandlePropertyChanged"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object that can be registered as an INotifyPropertyChanged handler."
'@Folder MVVM.Infrastructure.Bindings
'@ModuleDescription "An object that can be registered as an INotifyPropertyChanged handler."
'@Interface
'@Exposed
Option Explicit

'@Description "A callback invoked when a property value is set."
Public Sub OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
Attribute OnPropertyChanged.VB_Description = "A callback invoked when a property value is set."
End Sub
