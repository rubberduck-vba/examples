VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITypeMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an object responsible for mapping programmatic data types to ADODB.DataTypeEnum values."
'@Folder("SecureADODB.ParameterProvider.Abstract")
'@ModuleDescription("Represents an object responsible for mapping programmatic data types to ADODB.DataTypeEnum values.")
'@Interface
'@Exposed
Option Explicit

'@Description("Gets/sets the ADODB.DataTypeEnum mapping for the specified data type name.")
Public Property Get Mapping(ByVal name As String) As ADODB.DataTypeEnum
Attribute Mapping.VB_Description = "Gets/sets the ADODB.DataTypeEnum mapping for the specified data type name."
End Property

Public Property Let Mapping(ByVal name As String, ByVal value As ADODB.DataTypeEnum)
End Property

'@Description("True if the specified data type name is mapped to an ADODB.DataTypeEnum value.")
Public Property Get IsMapped(ByVal name As String) As Boolean
Attribute IsMapped.VB_Description = "True if the specified data type name is mapped to an ADODB.DataTypeEnum value."
End Property
