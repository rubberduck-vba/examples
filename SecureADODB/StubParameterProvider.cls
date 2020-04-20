VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StubParameterProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A stub acting as a IParameterProvider implementation."
'@Folder("Tests.Stubs")
'@ModuleDescription("A stub acting as a IParameterProvider implementation.")
Option Explicit
Implements IParameterProvider

Private Function IParameterProvider_FromValue(ByVal value As Variant, Optional ByVal name As String, Optional ByVal dataTypeName As String) As ADODB.Parameter
    Set IParameterProvider_FromValue = New ADODB.Parameter
End Function

Private Function IParameterProvider_FromValues(values() As Variant) As Collection
    Dim result As VBA.Collection
    Set result = New VBA.Collection
    Set IParameterProvider_FromValues = result
End Function
