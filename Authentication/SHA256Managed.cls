VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SHA256Managed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("Authentication.Hashing")
'@PredeclaredId
Option Explicit
Implements IHashAlgorithm

Private base As HashAlgorithmBase

'@Description("Factory method creates and returns a new instance of this class.")
Public Function Create() As IHashAlgorithm
    Set Create = New SHA256Managed
End Function

Private Sub Class_Initialize()
    Set base = New HashAlgorithmBase
End Sub

Private Function IHashAlgorithm_ComputeHash(ByVal value As String) As String
    Dim bytes() As Byte
    bytes = StrConv(value, vbFromUnicode)
    
    Dim algo As Object
    Set algo = CreateObject("System.Security.Cryptography.SHA256Managed")
    
    Dim buffer() As Byte
    buffer = algo.ComputeHash_2(bytes)
        
    IHashAlgorithm_ComputeHash = base.ToHexString(buffer)
End Function
