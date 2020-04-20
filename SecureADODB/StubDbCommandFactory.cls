VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StubDbCommandFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A stub acting as a IDbCommandFactory implementation that outputs StubDbCommand objects."
'@Folder("Tests.Stubs")
'@ModuleDescription("A stub acting as a IDbCommandFactory implementation that outputs StubDbCommand objects.")
Option Explicit
Implements IDbCommandFactory
Private Type TInvokeState
    CreateCommandInvokes As Long
End Type
Private this As TInvokeState

Public Property Get CreateCommandInvokes() As Long
    CreateCommandInvokes = this.CreateCommandInvokes
End Property

Private Function IDbCommandFactory_Create(ByVal db As IDbConnection) As IDbCommand
    this.CreateCommandInvokes = this.CreateCommandInvokes + 1
    Set IDbCommandFactory_Create = New StubDbCommand
End Function
