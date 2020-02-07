VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAuthPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("Authentication.Abstract")
'@Interface
Option Explicit

'@Description("True if user is authenticated")
Public Property Get IsAuthenticated() As Boolean
End Property

'@Description("Prompts for user credentials. Returns False if prompt is cancelled.")
Public Function Authenticate() As Boolean
End Function
