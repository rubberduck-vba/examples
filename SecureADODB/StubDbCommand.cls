VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StubDbCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A stub acting as a IDbCommand implementation."
'@Folder("Tests.Stubs")
'@ModuleDescription("A stub acting as a IDbCommand implementation.")
'@IgnoreModule ProcedureNotUsed
Option Explicit
Implements IDbCommand
Private Type TInvokeState
    ExecuteInvokes As Long
    ExecuteNonQueryInvokes As Long
    ExecuteWithParameterInvokes As Long
    GetSingleValueInvokes As Long
End Type
Private this As TInvokeState

Public Property Get ExecuteInvokes() As Long
    ExecuteInvokes = this.ExecuteInvokes
End Property

Public Property Get ExecuteNonQueryInvokes() As Long
    ExecuteNonQueryInvokes = this.ExecuteNonQueryInvokes
End Property

Public Property Get ExecuteWithParameterInvokes() As Long
    ExecuteWithParameterInvokes = this.ExecuteWithParameterInvokes
End Property

Public Property Get GetSingleValueInvokes() As Long
    GetSingleValueInvokes = this.GetSingleValueInvokes + 1
End Property

Private Function IDbCommand_Execute(ByVal sql As String, ParamArray args() As Variant) As ADODB.Recordset
    this.ExecuteInvokes = this.ExecuteInvokes + 1
    Set IDbCommand_Execute = New ADODB.Recordset
End Function

Private Sub IDbCommand_ExecuteNonQuery(ByVal sql As String, ParamArray args() As Variant)
    this.ExecuteNonQueryInvokes = this.ExecuteNonQueryInvokes + 1
End Sub

Private Function IDbCommand_ExecuteWithParameters(ByVal sql As String, ParamArray params() As Variant) As ADODB.Recordset
    this.ExecuteWithParameterInvokes = this.ExecuteWithParameterInvokes + 1
    Set IDbCommand_ExecuteWithParameters = New ADODB.Recordset
End Function

Private Function IDbCommand_GetSingleValue(ByVal sql As String, ParamArray args() As Variant) As Variant
    this.GetSingleValueInvokes = this.GetSingleValueInvokes + 1
    IDbCommand_GetSingleValue = Empty
End Function
