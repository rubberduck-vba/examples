Attribute VB_Name = "ParameterProviderTests"
'@TestModule
'@Folder("Tests")
'@IgnoreModule
Option Explicit
Option Private Module

Private Const ExpectedError As Long = SecureADODBCustomError
#Const LateBind = LateBindTests

#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If

'@ModuleInitialize
Private Sub ModuleInitialize()
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
    #Else
        Set Assert = New Rubberduck.PermissiveAssertClass
    #End If
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

Private Function GetSUT() As IParameterProvider
    Set GetSUT = AdoParameterProvider.Create(GetDefaultMappings)
End Function

Private Function GetDefaultMappings() As ITypeMap
    Set GetDefaultMappings = AdoTypeMappings.Default
End Function

'@TestMethod("Factory Guard")
Private Sub Create_ThrowsIfNotInvokedFromDefaultInstance()
    On Error GoTo TestFail
    
    With New AdoParameterProvider
        On Error GoTo CleanFail
        Dim sut As AdoParameterProvider
        Set sut = .Create(GetDefaultMappings)
        On Error GoTo 0
    End With
    
CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("Factory Guard")
Private Sub Create_ThrowsGivenNullMappings()
    
    On Error GoTo CleanFail
    Dim sut As IParameterProvider
    Set sut = AdoParameterProvider.Create(Nothing)
    On Error GoTo 0

CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("Guard Clauses")
Private Sub TypeMappings_ThrowsIfAlreadySet()
    On Error GoTo TestFail
    
    Dim sut As AdoParameterProvider
    Set sut = AdoParameterProvider.Create(GetDefaultMappings)
    
    On Error GoTo CleanFail
    Set sut.TypeMappings = GetDefaultMappings
    On Error GoTo 0
    
CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("Guard Clauses")
Private Sub TypeMappings_ThrowsGivenNullMappings()
    
    On Error GoTo CleanFail
    Dim sut As AdoParameterProvider
    Set sut = AdoParameterProvider.Create(Nothing)
    On Error GoTo 0
    
CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("ParameterProvider")
Private Sub FromValue_MapsParameterSizeToStringLength()
    Dim sut As IParameterProvider
    Set sut = GetSUT
    
    Const value = "ABC XYZ"
    
    Dim p As ADODB.Parameter
    Set p = sut.FromValue(value)
    
    Assert.AreEqual Len(value), p.Size
End Sub

'@TestMethod("ParameterProvider")
Private Sub FromValue_MapsParameterTypeAsPerMapping()
    Const expected = DataTypeEnum.adNumeric
    Const value = 42

    Dim typeMap As ITypeMap
    Set typeMap = AdoTypeMappings.Default()
    If typeMap.Mapping(TypeName(value)) = expected Then Assert.Inconclusive "'expected' data type should not be the default mapping for the specified 'value'."
    typeMap.Mapping(TypeName(value)) = expected

    Dim sut As IParameterProvider
    Set sut = AdoParameterProvider.Create(typeMap)
    
    Dim p As ADODB.Parameter
    Set p = sut.FromValue(value)
    
    Assert.AreEqual expected, p.Type
End Sub

'@TestMethod("ParameterProvider")
Private Sub FromValue_CreatesInputParameters()
    Const expected = ADODB.adParamInput
    Const value = 42
    
    Dim sut As IParameterProvider
    Set sut = GetSUT
    
    Dim p As ADODB.Parameter
    Set p = sut.FromValue(value)
    
    Assert.AreEqual expected, p.direction
End Sub

'@TestMethod("ParameterProvider")
Private Sub FromValues_YieldsAsManyParametersAsSuppliedArgs()
    Dim sut As IParameterProvider
    Set sut = GetSUT
    
    Dim args(1 To 4) As Variant '1-based to match collection indexing
    args(1) = True
    args(2) = 42
    args(3) = 34567
    args(4) = "some string"
    
    Dim values As VBA.Collection
    Set values = sut.FromValues(args)
    
    Assert.AreEqual UBound(args), values.Count
End Sub

