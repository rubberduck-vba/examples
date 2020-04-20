VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StubDbCommandBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A stub acting as a IDbCommandBase implementation."
'@Folder("Tests.Stubs")
'@ModuleDescription("A stub acting as a IDbCommandBase implementation.")
'@IgnoreModule ProcedureNotUsed
Option Explicit
Implements IDbCommandBase
Private Type TInvokeState
    CreateCommandInvokes As Long
    GetDisconnectedRecordsetInvokes As Long
    GetSingleValueInvokes As Long
    ValidateOrdinalArgumentsInvokes As Long
    SingleValueResult As Variant
End Type
Private this As TInvokeState

Public Property Get CreateCommandInvokes() As Long
    CreateCommandInvokes = this.CreateCommandInvokes
End Property

Public Property Get GetDisconnectedRecordsetInvokes() As Long
    GetDisconnectedRecordsetInvokes = this.GetDisconnectedRecordsetInvokes
End Property

Public Property Get GetSingleValueInvokes() As Long
    GetSingleValueInvokes = this.GetSingleValueInvokes
End Property

Public Property Get ValidateOrdinalArgumentsInvokes() As Long
    ValidateOrdinalArgumentsInvokes = this.ValidateOrdinalArgumentsInvokes
End Property

Public Property Get SingleValueResult() As Variant
    SingleValueResult = this.SingleValueResult
End Property

Public Property Let SingleValueResult(ByVal value As Variant)
    this.SingleValueResult = value
End Property

Private Function IDbCommandBase_CreateCommand(ByVal db As IDbConnection, ByVal commandType As ADODB.CommandTypeEnum, ByVal sql As String, args() As Variant) As ADODB.Command
    this.CreateCommandInvokes = this.CreateCommandInvokes + 1
    Set IDbCommandBase_CreateCommand = New ADODB.Command
End Function

Private Function IDbCommandBase_GetDisconnectedRecordset(ByVal cmd As ADODB.Command) As ADODB.Recordset
    this.GetDisconnectedRecordsetInvokes = this.GetDisconnectedRecordsetInvokes + 1
    Set IDbCommandBase_GetDisconnectedRecordset = New ADODB.Recordset
End Function

Private Function IDbCommandBase_GetSingleValue(ByVal db As IDbConnection, ByVal sql As String, args() As Variant) As Variant
    this.GetSingleValueInvokes = this.GetSingleValueInvokes + 1
    IDbCommandBase_GetSingleValue = this.SingleValueResult
End Function

Private Function IDbCommandBase_ValidateOrdinalArguments(ByVal sql As String, args() As Variant) As Boolean
    this.ValidateOrdinalArgumentsInvokes = this.ValidateOrdinalArgumentsInvokes + 1
    IDbCommandBase_ValidateOrdinalArguments = DbCommandBase.ValidateOrdinalArguments(sql, args)
End Function
