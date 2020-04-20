VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StubDbConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A stub acting as a IDbConnection implementation."
'@Folder("Tests.Stubs")
'@ModuleDescription("A stub acting as a IDbConnection implementation.")
'@IgnoreModule ProcedureNotUsed
Option Explicit
Implements IDbConnection
Private stubState As ADODB.ObjectStateEnum
Private stubConnection As ADODB.Connection
Private beganTransaction As Boolean
Private committedTransaction As Boolean
Private rolledBackTransaction As Boolean

Public Property Get State() As ADODB.ObjectStateEnum
    State = stubState
End Property

Public Property Let State(ByVal value As ADODB.ObjectStateEnum)
    stubState = value
End Property

Public Property Get DidBeginTransaction() As Boolean
    DidBeginTransaction = beganTransaction
End Property

Public Property Get DidCommitTransaction() As Boolean
    DidCommitTransaction = committedTransaction
End Property

Public Property Get DidRollBackTransaction() As Boolean
    DidRollBackTransaction = rolledBackTransaction
End Property

Private Sub Class_Initialize()
    'needed because we don't want IDbConnection.AdoConnection to return Nothing.
    Set stubConnection = New ADODB.Connection
    stubState = adStateOpen
End Sub

Private Sub Class_Terminate()
    stubState = adStateClosed
End Sub

Private Property Get IDbConnection_AdoConnection() As ADODB.Connection
    Set IDbConnection_AdoConnection = stubConnection
End Property

Private Sub IDbConnection_BeginTransaction()
    beganTransaction = True
End Sub

Private Sub IDbConnection_CommitTransaction()
    committedTransaction = True
End Sub

Private Function IDbConnection_CreateCommand(ByVal commandType As ADODB.CommandTypeEnum, ByVal sql As String) As ADODB.Command
    Dim result As ADODB.Command
    Set result = New ADODB.Command
    'Set result.ActiveConnection = stubConnection '<~ we don't want to do that in tests.
    result.commandType = commandType
    result.CommandText = sql
    Set IDbConnection_CreateCommand = result
End Function

Private Sub IDbConnection_RollbackTransaction()
    rolledBackTransaction = True
End Sub

Private Property Get IDbConnection_Self() As IDbConnection
    Set IDbConnection_Self = Me
End Property

Private Property Get IDbConnection_State() As ADODB.ObjectStateEnum
    IDbConnection_State = State
End Property
