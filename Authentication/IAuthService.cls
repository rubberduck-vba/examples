VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAuthService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("Authentication.Abstract")
'@ModuleDescription("An interface representing an authentication mechanism.")
'@Interface
Option Explicit

'@Description("True if the supplied credentials are valid, False otherwise.")
Public Function Authenticate(ByVal model As UserAuthModel) As Boolean
End Function
