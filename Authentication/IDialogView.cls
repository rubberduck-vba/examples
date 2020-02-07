VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDialogView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("Authentication.Abstract")
'@Interface
Option Explicit

'@Description("Shows the view as a modal dialog. Returns True unless the dialog is cancelled.")
Public Function ShowDialog() As Boolean
End Function

Public Property Get ViewModel() As Object
End Property
