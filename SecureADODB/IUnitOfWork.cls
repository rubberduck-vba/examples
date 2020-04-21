VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IUnitOfWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an object encapsulating a database transaction."
'@Folder("SecureADODB.UnitOfWork")
'@ModuleDescription("Represents an object encapsulating a database transaction.")
'@Interface
'@Exposed
Option Explicit

'@Description("Commits the transaction.")
Public Sub Commit()
Attribute Commit.VB_Description = "Commits the transaction."
End Sub

'@Description("Rolls back the transaction.")
Public Sub Rollback()
Attribute Rollback.VB_Description = "Rolls back the transaction."
End Sub

'@Description("Creates a new command to execute as part of the transaction.")
'@Ignore ShadowedDeclaration: false positive here,
Public Function Command() As IDbCommand
Attribute Command.VB_Description = "Creates a new command to execute as part of the transaction."
End Function
