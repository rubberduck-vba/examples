Attribute VB_Name = "AutoDbCommandTests"
'@TestModule
'@Folder("Tests")
'@IgnoreModule

Option Explicit
Option Private Module

Private Const ERR_INVALID_WITHOUT_LIVE_CONNECTION As Long = 3709 ' raised by ADODB
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

Private Function GetSUT(Optional ByRef stubBase As StubDbCommandBase, Optional ByRef stubFactory As StubDbConnectionFactory) As IDbCommand
    Set stubFactory = New StubDbConnectionFactory
    Set stubBase = New StubDbCommandBase
    
    Dim result As AutoDbCommand
    Set result = AutoDbCommand.Create("connection string", stubFactory, stubBase)
    
    Set GetSUT = result
End Function

Private Function GetSingleParameterSelectSql() As String
    GetSingleParameterSelectSql = "SELECT * FROM [dbo].[Table1] WHERE [Field1] = ?;"
End Function

Private Function GetTwoParameterSelectSql() As String
    GetTwoParameterSelectSql = "SELECT * FROM [dbo].[Table1] WHERE [Field1] = ? AND [Field2] = ?;"
End Function

Private Function GetSingleParameterInsertSql() As String
    GetSingleParameterInsertSql = "INSERT INTO [dbo].[Table1] ([Timestamp], [Value]) VALUES (GETDATE(), ?);"
End Function

Private Function GetTwoParameterInsertSql() As String
    GetTwoParameterInsertSql = "INSERT INTO [dbo].[Table1] ([Timestamp], [Value], [ThingID]) VALUES (GETDATE(), ?, ?);"
End Function

Private Function GetStubParameter() As ADODB.Parameter
    Dim stubParameter As ADODB.Parameter
    Set stubParameter = New ADODB.Parameter
    stubParameter.value = 42
    stubParameter.Type = adInteger
    stubParameter.direction = adParamInput
    Set GetStubParameter = stubParameter
End Function

'@TestMethod("Factory Guard")
Private Sub Create_ThrowsIfNotInvokedFromDefaultInstance()
    On Error GoTo TestFail
    
    With New AutoDbCommand
        On Error GoTo CleanFail
        Dim sut As IDbCommand
        Set sut = .Create("connection string", New StubDbConnectionFactory, New StubDbCommandBase)
        On Error GoTo 0
    End With
    
CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("Factory Guard")
Private Sub Create_ThrowsGivenEmptyConnectionString()
    
    On Error GoTo CleanFail
    Dim sut As IDbCommand
    Set sut = AutoDbCommand.Create(vbNullString, New StubDbConnectionFactory, New StubDbCommandBase)
    On Error GoTo 0
    
CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("Factory Guard")
Private Sub Create_ThrowsGivenNullConnectionFactory()
    
    On Error GoTo CleanFail
    Dim sut As IDbCommand
    Set sut = AutoDbCommand.Create("connection string", Nothing, New StubDbCommandBase)
    On Error GoTo 0
    
CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("Factory Guard")
Private Sub Create_ThrowsGivenNullCommandBase()
    
    On Error GoTo CleanFail
    Dim sut As IDbCommand
    Set sut = AutoDbCommand.Create("connection string", New StubDbConnectionFactory, Nothing)
    On Error GoTo 0
    
CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("Guard Clauses")
Private Sub ConnectionFactory_ThrowsIfAlreadySet()
    On Error GoTo TestFail
    
    Dim sut As AutoDbCommand
    Set sut = AutoDbCommand.Create("connection string", New StubDbConnectionFactory, New StubDbCommandBase)
    
    On Error GoTo CleanFail
    Set sut.ConnectionFactory = New StubDbConnectionFactory
    On Error GoTo 0
    
CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("Guard Clauses")
Private Sub Base_ThrowsIfAlreadySet()
    On Error GoTo TestFail
    
    Dim sut As AutoDbCommand
    Set sut = AutoDbCommand.Create("connection string", New StubDbConnectionFactory, New StubDbCommandBase)
    
    On Error GoTo CleanFail
    Set sut.Base = New StubDbCommandBase
    On Error GoTo 0
    
CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("AutoDbCommand")
Private Sub Execute_ThrowsGivenExtraneousArgument()
    On Error GoTo TestFail
    
    Dim stubBase As StubDbCommandBase
    Dim stubFactory As StubDbConnectionFactory
    
    Dim sut As IDbCommand
    Set sut = GetSUT(stubBase, stubFactory)
    
    On Error GoTo CleanFail
    Dim result As ADODB.Recordset
    Set result = sut.Execute(GetSingleParameterSelectSql, 42, "extraneous argument")
    On Error GoTo 0
    
CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("AutoDbCommand")
Private Sub Execute_ThrowsGivenMissingArgument()
    On Error GoTo TestFail
    
    Dim stubBase As StubDbCommandBase
    Dim stubFactory As StubDbConnectionFactory
    
    Dim sut As IDbCommand
    Set sut = GetSUT(stubBase, stubFactory)
    
    On Error GoTo CleanFail
    Dim result As ADODB.Recordset
    Set result = sut.Execute(GetTwoParameterSelectSql, 42)
    On Error GoTo 0
    
CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("AutoDbCommand")
Private Sub ExecuteWithParameters_ThrowsGivenExtraneousArgument()
    On Error GoTo TestFail
    
    Dim stubBase As StubDbCommandBase
    Dim stubFactory As StubDbConnectionFactory
    
    Dim sut As IDbCommand
    Set sut = GetSUT(stubBase, stubFactory)
    
    On Error GoTo CleanFail
    Dim result As ADODB.Recordset
    Set result = sut.ExecuteWithParameters(GetSingleParameterSelectSql, GetStubParameter, GetStubParameter)
    On Error GoTo 0
    
CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("AutoDbCommand")
Private Sub ExecuteWithParameters_ThrowsGivenMissingArgument()
    On Error GoTo TestFail
    
    Dim stubBase As StubDbCommandBase
    Dim stubFactory As StubDbConnectionFactory
    
    Dim sut As IDbCommand
    Set sut = GetSUT(stubBase, stubFactory)
    
    On Error GoTo CleanFail
    Dim result As ADODB.Recordset
    Set result = sut.ExecuteWithParameters(GetTwoParameterSelectSql, GetStubParameter)
    On Error GoTo 0
    
CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("AutoDbCommand")
Private Sub ExecuteNonQuery_ThrowsGivenExtraneousArgument()
    On Error GoTo TestFail
    
    Dim stubBase As StubDbCommandBase
    Dim stubFactory As StubDbConnectionFactory
    
    Dim sut As IDbCommand
    Set sut = GetSUT(stubBase, stubFactory)
    
    On Error GoTo CleanFail
    sut.ExecuteNonQuery GetSingleParameterInsertSql, 42, "extraneous argument"
    On Error GoTo 0
    
CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("AutoDbCommand")
Private Sub ExecuteNonQuery_ThrowsGivenMissingArgument()
    On Error GoTo TestFail
    
    Dim stubBase As StubDbCommandBase
    Dim stubFactory As StubDbConnectionFactory
    
    Dim sut As IDbCommand
    Set sut = GetSUT(stubBase, stubFactory)
    
    On Error GoTo CleanFail
    sut.ExecuteNonQuery GetTwoParameterInsertSql, 42
    On Error GoTo 0

CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("AutoDbCommand")
Private Sub GetSingleValue_ThrowsGivenExtraneousArgument()
    On Error GoTo TestFail
    
    Dim stubBase As StubDbCommandBase
    Dim stubFactory As StubDbConnectionFactory
    
    Dim sut As IDbCommand
    Set sut = GetSUT(stubBase, stubFactory)
    
    On Error GoTo CleanFail
    Dim result As Variant
    result = sut.GetSingleValue(GetSingleParameterSelectSql, 42, "extraneous argument")
    On Error GoTo 0

CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("AutoDbCommand")
Private Sub GetSingleValue_ThrowsGivenMissingArgument()
    On Error GoTo TestFail
    
    Dim stubBase As StubDbCommandBase
    Dim stubFactory As StubDbConnectionFactory
    
    Dim sut As IDbCommand
    Set sut = GetSUT(stubBase, stubFactory)
    
    On Error GoTo CleanFail
    Dim result As Variant
    result = sut.GetSingleValue(GetTwoParameterSelectSql, 42)
    On Error GoTo 0

CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("AutoDbCommand")
Private Sub Execute_CreatesDbConnection()
    Dim stubBase As StubDbCommandBase
    Dim stubFactory As StubDbConnectionFactory
    
    Dim sut As IDbCommand
    Set sut = GetSUT(stubBase, stubFactory)
    
    Dim result As ADODB.Recordset
    Set result = sut.Execute(GetSingleParameterSelectSql, 42)
    
    Assert.AreEqual 1, stubFactory.CreateConnectionInvokes
End Sub

'@TestMethod("AutoDbCommand")
Private Sub ExecuteNonQuery_CreatesDbConnection()
    Dim stubBase As StubDbCommandBase
    Dim stubFactory As StubDbConnectionFactory
    
    Dim sut As IDbCommand
    Set sut = GetSUT(stubBase, stubFactory)
    
    On Error Resume Next
    sut.ExecuteNonQuery GetSingleParameterInsertSql, 42
    Debug.Assert Err.Number = ERR_INVALID_WITHOUT_LIVE_CONNECTION
    On Error GoTo 0
    
    Assert.AreEqual 1, stubFactory.CreateConnectionInvokes
End Sub

'@TestMethod("AutoDbCommand")
Private Sub ExecuteWithParameters_CreatesDbConnection()
    Dim stubBase As StubDbCommandBase
    Dim stubFactory As StubDbConnectionFactory
    
    Dim sut As IDbCommand
    Set sut = GetSUT(stubBase, stubFactory)
    
    Dim result As Recordset
    Set result = sut.ExecuteWithParameters(GetTwoParameterInsertSql, GetStubParameter, GetStubParameter)
    
    Assert.AreEqual 1, stubFactory.CreateConnectionInvokes
End Sub

'@TestMethod("AutoDbCommand")
Private Sub GetSingleValue_CreatesDbConnection()
    Dim stubBase As StubDbCommandBase
    Dim stubFactory As StubDbConnectionFactory
    
    Dim sut As IDbCommand
    Set sut = GetSUT(stubBase, stubFactory)
    
    Dim result As Variant
    result = sut.GetSingleValue(GetSingleParameterSelectSql, 42)
    
    Assert.AreEqual 1, stubFactory.CreateConnectionInvokes
End Sub

'@TestMethod("AutoDbCommand")
Private Sub Execute_ReturnsDisconnectedRecordset()
    Dim stubBase As StubDbCommandBase
    Dim stubFactory As StubDbConnectionFactory
    
    Dim sut As IDbCommand
    Set sut = GetSUT(stubBase, stubFactory)
    
    Dim result As ADODB.Recordset
    Set result = sut.Execute(GetSingleParameterSelectSql, 42)
    
    Assert.AreEqual 1, stubBase.GetDisconnectedRecordsetInvokes
End Sub

'@TestMethod("AutoDbCommand")
Private Sub ExecuteWithParameters_ReturnsDisconnectedRecordset()
    Dim stubBase As StubDbCommandBase
    Dim stubFactory As StubDbConnectionFactory
    
    Dim sut As IDbCommand
    Set sut = GetSUT(stubBase, stubFactory)
    
    Dim result As ADODB.Recordset
    Set result = sut.ExecuteWithParameters(GetSingleParameterSelectSql, GetStubParameter)

    Assert.AreEqual 1, stubBase.GetDisconnectedRecordsetInvokes
End Sub

'@TestMethod("Validation")
Private Sub Execute_ValidatesArguments()
    Dim stubBase As StubDbCommandBase
    Dim stubFactory As StubDbConnectionFactory
    
    Dim sut As IDbCommand
    Set sut = GetSUT(stubBase, stubFactory)
    
    Dim result As ADODB.Recordset
    Set result = sut.Execute(GetSingleParameterSelectSql, 42)
    
    Assert.AreEqual 1, stubBase.ValidateOrdinalArgumentsInvokes
End Sub

'@TestMethod("Validation")
Private Sub ExecuteNonQuery_ValidatesArguments()
    Dim stubBase As StubDbCommandBase
    Dim stubFactory As StubDbConnectionFactory
    
    Dim sut As IDbCommand
    Set sut = GetSUT(stubBase, stubFactory)
    
    On Error Resume Next
    sut.ExecuteNonQuery GetSingleParameterInsertSql, 42
    Debug.Assert Err.Number = ERR_INVALID_WITHOUT_LIVE_CONNECTION
    On Error GoTo 0
    
    Assert.AreEqual 1, stubBase.ValidateOrdinalArgumentsInvokes
End Sub

'@TestMethod("Validation")
Private Sub GetSingleValue_ValidatesArguments()
    Dim stubBase As StubDbCommandBase
    Dim stubFactory As StubDbConnectionFactory
    
    Dim sut As IDbCommand
    Set sut = GetSUT(stubBase, stubFactory)
    
    Dim result As Variant
    result = sut.GetSingleValue(GetSingleParameterSelectSql, 42)
    
    Assert.AreEqual 1, stubBase.ValidateOrdinalArgumentsInvokes
End Sub

'@TestMethod("Validation")
Private Sub ExecuteWithParameters_ValidatesArguments()
    Dim stubBase As StubDbCommandBase
    Dim stubFactory As StubDbConnectionFactory
    
    Dim sut As IDbCommand
    Set sut = GetSUT(stubBase, stubFactory)
    
    Dim result As ADODB.Recordset
    Set result = sut.ExecuteWithParameters(GetSingleParameterInsertSql, GetStubParameter)
    
    Assert.AreEqual 1, stubBase.ValidateOrdinalArgumentsInvokes
End Sub

