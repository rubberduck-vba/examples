VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HashAlgorithmBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Authentication.Hashing")
'@ModuleDescription("Provides common functionality used by IHashAlgorithm implementations.")
Option Explicit

'@Description("Converts a byte array to a string representation.")
Public Function ToHexString(ByRef buffer() As Byte) As String
    Dim result As String

    Dim i As Long
    For i = LBound(buffer) To UBound(buffer)
        result = result & Hex(buffer(i))
    Next

    ToHexString = result
End Function
