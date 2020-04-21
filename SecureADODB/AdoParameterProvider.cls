VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdoParameterProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for converting a value of a mapped data type to an ADODB.Parameter object."
'@Exposed
'@Folder("SecureADODB.ParameterProvider")
'@ModuleDescription("An object responsible for converting a value of a mapped data type to an ADODB.Parameter object.")
'@PredeclaredId
Option Explicit
Implements IParameterProvider
Private mappings As ITypeMap

'@Inject: just an idea.. see #https://github.com/rubberduck-vba/Rubberduck/issues/5463
'@Description("Creates a new parameter provider using the specified type mappings.")
Public Function Create(ByVal typeMap As ITypeMap) As IParameterProvider
Attribute Create.VB_Description = "Creates a new parameter provider using the specified type mappings."
    Errors.GuardNonDefaultInstance Me, AdoParameterProvider
    Errors.GuardNullReference typeMap
    
    Dim result As AdoParameterProvider
    Set result = New AdoParameterProvider
    
    Set result.TypeMappings = typeMap
    Set Create = result

End Function

'@Inject: this member should only be invoked by Me.Create, where Me is the class' default/predeclared instance.
'@Description "Gets/sets the type mappings for this instance."
'@Ignore ProcedureNotUsed: false positive with v2.5.0.5418
Friend Property Set TypeMappings(ByVal value As ITypeMap)
Attribute TypeMappings.VB_Description = "Gets/sets the type mappings for this instance."
    Errors.GuardNullReference value
    Errors.GuardDoubleInitialization mappings
    
    Set mappings = value
    
End Property

Private Function IParameterProvider_FromValue(ByVal value As Variant, Optional ByVal name As String, Optional ByVal dataTypeName As String) As ADODB.Parameter
    If TypeOf value Is ADODB.Parameter Then
        Set IParameterProvider_FromValue = value
        Exit Function
    End If

    Dim localDataTypeName As String
    localDataTypeName = dataTypeName
    If localDataTypeName = vbNullString Then localDataTypeName = TypeName(value)
    Errors.GuardExpression Not mappings.IsMapped(localDataTypeName), message:="The data type '" & localDataTypeName & "' has no ADODB.DataTypeEnum mapping."
    
    Dim result As ADODB.Parameter
    Set result = New ADODB.Parameter
    result.Type = mappings.Mapping(localDataTypeName)
    result.direction = ADODB.adParamInput
    
    If result.Type = AdoTypeMappings.DefaultNullMapping Then
        result.Attributes = ParameterAttributesEnum.adParamNullable
        result.value = Null
    Else
        result.value = value
    End If
    
    If IsCharMapping(result.Type) And Len(value) > 0 Then result.Size = Len(value)
    If name <> vbNullString Then result.name = name
    
    Set IParameterProvider_FromValue = result
End Function

Private Function IsCharMapping(ByVal value As ADODB.DataTypeEnum) As Boolean
    IsCharMapping = value = adChar _
                           Or value = adWChar _
                           Or value = adVarChar _
                           Or value = adVarWChar
End Function

Private Function IParameterProvider_FromValues(ByRef values() As Variant) As VBA.Collection
    Dim results As VBA.Collection
    Set results = New VBA.Collection
    On Error GoTo CleanFail
    Dim i As Long
    For i = LBound(values) To UBound(values)
        results.Add IParameterProvider_FromValue(values(i))
    Next
CleanExit:
    Set IParameterProvider_FromValues = results
    Exit Function
CleanFail:
    If Err.Number <> 9 Then
        Errors.RethrowOnError
    End If
    Resume CleanExit
End Function
