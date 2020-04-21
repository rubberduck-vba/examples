VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A thin wrapper around an ADODB connection."
'@Exposed
'@Folder("SecureADODB.DbConnection")
'@ModuleDescription("A thin wrapper around an ADODB connection.")
'@PredeclaredId
Option Explicit
Implements IDbConnection
'@Ignore MissingMemberAnnotation
Private WithEvents AdoConnection As ADODB.Connection
Attribute AdoConnection.VB_VarHelpID = -1
Private HasActiveTransaction As Boolean

'@Inject: just an idea.. see #https://github.com/rubberduck-vba/Rubberduck/issues/5463
Public Function Create(ByVal connString As String) As IDbConnection
    Errors.GuardNonDefaultInstance Me, DbConnection
    
    Dim result As DbConnection
    Set result = New DbConnection
    
    result.Connect connString
    Set Create = result

End Function

'@Inject: this member should only be invoked by Me.Create, where Me is the class' default/predeclared instance.
Friend Sub Connect(ByVal connString As String)
    Errors.GuardExpression AdoConnection.ConnectionString <> vbNullString, message:="Object is already initialized."
    Errors.GuardEmptyString connString
    
    AdoConnection.Open connString

End Sub

Private Sub AdoConnection_BeginTransComplete(ByVal TransactionLevel As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Begin transaction completed. "
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.Number, pError.Description
End Sub

Private Sub AdoConnection_CommitTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Commit transaction completed. "
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.Number, pError.Description
End Sub

Private Sub AdoConnection_ConnectComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Connect completed. Status: " & adStatus
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.Number, pError.Description
End Sub

Private Sub AdoConnection_Disconnect(adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Disconnect completed. Status: " & adStatus
End Sub

Private Sub AdoConnection_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    Debug.Print "Execute completed, " & RecordsAffected & " record(s) affected."
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.Number, pError.Description
End Sub

Private Sub AdoConnection_InfoMessage(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    If pError.Number = 0 Then
        Debug.Print "Info: " & pError.Description
    Else
        Debug.Print "Error: " & pError.Number, pError.Description
    End If
End Sub

Private Sub AdoConnection_RollbackTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Rollback transaction completed."
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.Number, pError.Description
End Sub

Private Sub AdoConnection_WillConnect(ConnectionString As String, UserID As String, Password As String, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Begin connect..."
End Sub

Private Sub AdoConnection_WillExecute(Source As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    Debug.Print "Begin execute..."
End Sub

Private Sub Class_Initialize()
    Set AdoConnection = New ADODB.Connection
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    AdoConnection.Close
    On Error GoTo 0
    Set AdoConnection = Nothing
End Sub

Private Property Get IDbConnection_AdoConnection() As ADODB.Connection
    Set IDbConnection_AdoConnection = AdoConnection
End Property

Private Sub IDbConnection_BeginTransaction()
    On Error Resume Next 'not all providers support transactions
    Dim result As Long
    result = AdoConnection.BeginTrans
    On Error GoTo 0
    HasActiveTransaction = True
End Sub

Private Sub IDbConnection_CommitTransaction()
    On Error Resume Next
    AdoConnection.CommitTrans
    On Error GoTo 0
    HasActiveTransaction = False
End Sub

Private Function IDbConnection_CreateCommand(ByVal commandType As ADODB.CommandTypeEnum, ByVal sql As String) As ADODB.Command
    Errors.GuardEmptyString sql
    Errors.GuardExpression AdoConnection.State <> adStateOpen, message:="Connection is not open."
    
    Dim result As ADODB.Command
    Set result = New ADODB.Command
    
    Set result.ActiveConnection = AdoConnection
    result.commandType = commandType
    result.CommandText = sql
    
    Set IDbConnection_CreateCommand = result
End Function

Private Sub IDbConnection_RollbackTransaction()
    If HasActiveTransaction Then AdoConnection.RollbackTrans
    HasActiveTransaction = False
End Sub

Private Property Get IDbConnection_Self() As IDbConnection
    Set IDbConnection_Self = Me
End Property

Private Property Get IDbConnection_State() As ADODB.ObjectStateEnum
    IDbConnection_State = AdoConnection.State
End Property
