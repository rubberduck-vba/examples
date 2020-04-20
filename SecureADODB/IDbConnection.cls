VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an object that wraps an active ADODB connection."
'@Folder("SecureADODB.DbConnection.Abstract")
'@ModuleDescription("Represents an object that wraps an active ADODB connection.")
'@Exposed
'@Interface
Option Explicit

'@Description("Gets the wrapped ADODB connection.")
Public Property Get AdoConnection() As ADODB.Connection
Attribute AdoConnection.VB_Description = "Gets the wrapped ADODB connection."
End Property

'@Description("Gets the state of the wrapped ADODB connection.")
Public Property Get State() As ADODB.ObjectStateEnum
Attribute State.VB_Description = "Gets the state of the wrapped ADODB connection."
End Property

'@Description("Creates an ADODB.Command that uses the wrapped connection")
Public Function CreateCommand(ByVal commandType As ADODB.CommandTypeEnum, ByVal sql As String) As ADODB.Command
Attribute CreateCommand.VB_Description = "Creates an ADODB.Command that uses the wrapped connection"
End Function

'@Description("Returns the object itself. Useful to retrieve the With object variable in a With block.")
Public Property Get Self() As IDbConnection
Attribute Self.VB_Description = "Returns the object itself. Useful to retrieve the With object variable in a With block."
End Property

Public Sub BeginTransaction()
End Sub

Public Sub CommitTransaction()
End Sub

Public Sub RollbackTransaction()
End Sub
