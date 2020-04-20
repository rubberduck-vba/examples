Attribute VB_Name = "TypeMappingTests"
'@TestModule
'@Folder("Tests")
'@IgnoreModule
Option Explicit
Option Private Module

Private Const InvalidTypeName As String = "this isn't a valid type name"
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

'@TestMethod("Factory Guard")
Private Sub Default_ThrowsIfNotInvokedFromDefaultInstance()
    On Error GoTo TestFail
    With New AdoTypeMappings
        On Error GoTo CleanFail
        Dim sut As AdoTypeMappings
        Set sut = .Default
        On Error GoTo 0
    End With
CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

Private Sub DefaultMapping_MapsType(ByVal name As String)
    Dim sut As ITypeMap
    Set sut = AdoTypeMappings.Default
    Assert.IsTrue sut.IsMapped(name)
End Sub

'@TestMethod("Type Mappings")
Private Sub Mapping_ThrowsIfUndefined()
    On Error GoTo TestFail
    With AdoTypeMappings.Default
        On Error GoTo CleanFail
        Dim value As ADODB.DataTypeEnum
        value = .Mapping(InvalidTypeName)
        On Error GoTo 0
    End With
CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("Type Mappings")
Private Sub IsMapped_FalseIfUndefined()
    Dim sut As ITypeMap
    Set sut = AdoTypeMappings.Default
    Assert.IsFalse sut.IsMapped(InvalidTypeName)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsBoolean()
    Dim value As Boolean
    DefaultMapping_MapsType TypeName(value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsByte()
    Dim value As Byte
    DefaultMapping_MapsType TypeName(value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsCurrency()
    Dim value As Currency
    DefaultMapping_MapsType TypeName(value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsDate()
    Dim value As Date
    DefaultMapping_MapsType TypeName(value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsDouble()
    Dim value As Double
    DefaultMapping_MapsType TypeName(value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsInteger()
    Dim value As Integer
    DefaultMapping_MapsType TypeName(value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsLong()
    Dim value As Long
    DefaultMapping_MapsType TypeName(value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsSingle()
    Dim value As Single
    DefaultMapping_MapsType TypeName(value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsString()
    Dim value As String
    DefaultMapping_MapsType TypeName(value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsEmpty()
    Dim value As Variant
    DefaultMapping_MapsType TypeName(value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsNull()
    Dim value As Variant
    value = Null
    DefaultMapping_MapsType TypeName(value)
End Sub

Private Function GetDefaultMappingFor(ByVal name As String) As ADODB.DataTypeEnum
    On Error GoTo CleanFail
    Dim sut As ITypeMap
    Set sut = AdoTypeMappings.Default
    GetDefaultMappingFor = sut.Mapping(name)
    Exit Function
CleanFail:
    Assert.Inconclusive "Default mapping is undefined for '" & name & "'."
End Function

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForBoolean_MapsTo_adBoolean()
    Const expected = adBoolean
    Dim value As Boolean
    Assert.AreEqual expected, GetDefaultMappingFor(TypeName(value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForByte_MapsTo_adInteger()
    Const expected = adInteger
    Dim value As Byte
    Assert.AreEqual expected, GetDefaultMappingFor(TypeName(value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForCurrency_MapsTo_adCurrency()
    Const expected = adCurrency
    Dim value As Currency
    Assert.AreEqual expected, GetDefaultMappingFor(TypeName(value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForDate_MapsTo_adDate()
    Const expected = adDate
    Dim value As Date
    Assert.AreEqual expected, GetDefaultMappingFor(TypeName(value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForDouble_MapsTo_adDouble()
    Const expected = adDouble
    Dim value As Double
    Assert.AreEqual expected, GetDefaultMappingFor(TypeName(value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForInteger_MapsTo_adInteger()
    Const expected = adInteger
    Dim value As Integer
    Assert.AreEqual expected, GetDefaultMappingFor(TypeName(value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForLong_MapsTo_adInteger()
    Const expected = adInteger
    Dim value As Long
    Assert.AreEqual expected, GetDefaultMappingFor(TypeName(value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForNull_MapsTo_DefaultNullMapping()
    Dim expected As ADODB.DataTypeEnum
    expected = AdoTypeMappings.DefaultNullMapping
    Assert.AreEqual expected, GetDefaultMappingFor(TypeName(Null))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForEmpty_MapsTo_DefaultNullMapping()
    Dim expected As ADODB.DataTypeEnum
    expected = AdoTypeMappings.DefaultNullMapping
    Assert.AreEqual expected, GetDefaultMappingFor(TypeName(Empty))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForSingle_MapsTo_adSingle()
    Const expected = adSingle
    Dim value As Single
    Assert.AreEqual expected, GetDefaultMappingFor(TypeName(value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForString_MapsTo_adVarWChar()
    Const expected = adVarWChar
    Dim value As String
    Assert.AreEqual expected, GetDefaultMappingFor(TypeName(value))
End Sub

