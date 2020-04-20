VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IParameterProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an object responsible for converting a value of a mapped data type to an ADODB.Parameter object."
'@Folder("SecureADODB.ParameterProvider.Abstract")
'@ModuleDescription("Represents an object responsible for converting a value of a mapped data type to an ADODB.Parameter object.")
'@Interface
'@Exposed
Option Explicit

'@Description("Creates an ADODB.Parameter object from the specified value.")
Public Function FromValue(ByVal value As Variant, Optional ByVal name As String, Optional ByVal dataTypeName As String) As ADODB.Parameter
Attribute FromValue.VB_Description = "Creates an ADODB.Parameter object from the specified value."
End Function

'@Description("Creates an ADODB.Parameter object for each specified value; returns a VBA.Collection containing all parameters.")
Public Function FromValues(ByRef values() As Variant) As VBA.Collection
Attribute FromValues.VB_Description = "Creates an ADODB.Parameter object for each specified value; returns a VBA.Collection containing all parameters."
End Function

