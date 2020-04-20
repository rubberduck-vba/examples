VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DefaultDbCommandFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed
'@PredeclaredId
'@Folder("SecureADODB.DbCommand")
Option Explicit
Implements IDbCommandFactory
Private Type TFactory
    Base As IDbCommandBase
End Type
Private this As TFactory

'@Inject
Public Function Create(ByVal baseCommand As IDbCommandBase) As IDbCommandFactory
    Dim result As DefaultDbCommandFactory
    Set result = New DefaultDbCommandFactory
    Set result.Base = baseCommand
    Set Create = result
End Function

'@Inject: this member should only be invoked by Me.Create, where Me is the class' default/predeclared instance.
'@Ignore ProcedureNotUsed: false positive with v2.5.0.5418
Friend Property Set Base(ByVal value As IDbCommandBase)
    Set this.Base = value
End Property

Private Function IDbCommandFactory_Create(ByVal db As IDbConnection) As IDbCommand
    Set IDbCommandFactory_Create = DefaultDbCommand.Create(db, this.Base)
End Function

