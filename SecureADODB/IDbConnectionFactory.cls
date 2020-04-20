VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbConnectionFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An abstract factory that creates IDbConnection objects."
'@Folder("SecureADODB.DbConnection.Abstract")
'@ModuleDescription("An abstract factory that creates IDbConnection objects.")
'@Exposed
'@Interface
Option Explicit

'@Description("Creates and returns a new database connection.")
Public Function Create(ByVal connString As String) As IDbConnection
Attribute Create.VB_Description = "Creates and returns a new database connection."
End Function
