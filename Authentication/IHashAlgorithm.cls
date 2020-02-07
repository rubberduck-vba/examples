VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHashAlgorithm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("Authentication.Hashing")
'@ModuleDescription("An interface representing a hashing algorithm.")
'@Interface
Option Explicit

'@Description("Computes a hash for the given string value.")
Public Function ComputeHash(ByVal value As String) As String
End Function
