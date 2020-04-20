Attribute VB_Name = "UnitOfWorkTests"
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

'@TestMethod("Factory Guard")
Private Sub Create_ThrowsIfNotInvokedFromDefaultInstance()
    On Error GoTo TestFail
    
    With New UnitOfWork
        On Error GoTo CleanFail
        Dim sut As IUnitOfWork
        Set sut = .Create(New StubDbConnection, New StubDbCommandFactory)
        On Error GoTo 0
    End With
    
CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("Factory Guard")
Private Sub Create_ThrowsGivenNullConnection()
    
    On Error GoTo CleanFail
    Dim sut As IUnitOfWork
    Set sut = UnitOfWork.Create(Nothing, New StubDbCommandFactory)
    On Error GoTo 0
    
CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("Factory Guard")
Private Sub Create_ThrowsGivenConnectionStateNotOpen()
    On Error GoTo TestFail
    Dim db As StubDbConnection
    Set db = New StubDbConnection
    db.State = adStateClosed
    
    On Error GoTo CleanFail
    Dim sut As IUnitOfWork
    Set sut = UnitOfWork.Create(db, New StubDbCommandFactory)
    On Error GoTo 0
    
CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("Factory Guard")
Private Sub Create_ThrowsGivenNullCommandFactory()
    
    On Error GoTo CleanFail
    Dim sut As IUnitOfWork
    Set sut = UnitOfWork.Create(New StubDbConnection, Nothing)
    On Error GoTo 0
    
CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("Guard Clauses")
Private Sub CommandFactory_ThrowsIfAlreadySet()
    On Error GoTo TestFail
    
    Dim sut As UnitOfWork
    Set sut = UnitOfWork.Create(New StubDbConnection, New StubDbCommandFactory)
    
    On Error GoTo CleanFail
    Set sut.CommandFactory = New StubDbCommandFactory
    On Error GoTo 0
    
CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("Guard Clauses")
Private Sub Connection_ThrowsIfAlreadySet()
    On Error GoTo TestFail
    
    Dim sut As UnitOfWork
    Set sut = UnitOfWork.Create(New StubDbConnection, New StubDbCommandFactory)
    
    On Error GoTo CleanFail
    Set sut.Connection = New StubDbConnection
    On Error GoTo 0
    
CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("UnitOfWork")
Private Sub Command_CreatesDbCommandWithFactory()
    
    Dim stubCommandFactory As StubDbCommandFactory
    Set stubCommandFactory = New StubDbCommandFactory
    
    Dim sut As IUnitOfWork
    Set sut = UnitOfWork.Create(New StubDbConnection, stubCommandFactory)
    
    Dim result As IDbCommand
    Set result = sut.Command
    
    Assert.AreEqual 1, stubCommandFactory.CreateCommandInvokes
End Sub

'@TestMethod("UnitOfWork")
Private Sub Create_StartsTransaction()
    
    Dim stubConnection As StubDbConnection
    Set stubConnection = New StubDbConnection
    
    Dim sut As IUnitOfWork
    Set sut = UnitOfWork.Create(stubConnection, New StubDbCommandFactory)
    
    Assert.IsTrue stubConnection.DidBeginTransaction
End Sub

'@TestMethod("UnitOfWork")
Private Sub Commit_CommitsTransaction()
    
    Dim stubConnection As StubDbConnection
    Set stubConnection = New StubDbConnection
    
    Dim sut As IUnitOfWork
    Set sut = UnitOfWork.Create(stubConnection, New StubDbCommandFactory)
    
    sut.Commit
    
    Assert.IsTrue stubConnection.DidCommitTransaction
End Sub

'@TestMethod("UnitOfWork")
Private Sub Commit_ThrowsIfAlreadyCommitted()
    On Error GoTo TestFail
    
    Dim stubConnection As StubDbConnection
    Set stubConnection = New StubDbConnection
    
    Dim sut As IUnitOfWork
    Set sut = UnitOfWork.Create(stubConnection, New StubDbCommandFactory)
    
    sut.Commit
    On Error GoTo CleanFail
    sut.Commit
    On Error GoTo 0

CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("UnitOfWork")
Private Sub Commit_ThrowsIfAlreadyRolledBack()
    On Error GoTo TestFail
    
    Dim stubConnection As StubDbConnection
    Set stubConnection = New StubDbConnection
    
    Dim sut As IUnitOfWork
    Set sut = UnitOfWork.Create(stubConnection, New StubDbCommandFactory)
    
    sut.Rollback
    On Error GoTo CleanFail
    sut.Commit
    On Error GoTo 0

CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("UnitOfWork")
Private Sub Rollback_ThrowsIfAlreadyCommitted()
    On Error GoTo TestFail
    
    Dim stubConnection As StubDbConnection
    Set stubConnection = New StubDbConnection
    
    Dim sut As IUnitOfWork
    Set sut = UnitOfWork.Create(stubConnection, New StubDbCommandFactory)
    
    sut.Commit
    On Error GoTo CleanFail
    sut.Rollback
    On Error GoTo 0

CleanFail:
    If Err.Number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

