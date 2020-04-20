VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbCommandBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Defines an explicit public interface for the DbCommandBase class."
'@Exposed
'@Folder("SecureADODB.DbCommand.Abstract")
'@ModuleDescription("Defines an explicit public interface for the DbCommandBase class.")
'@Interface
Option Explicit

'@Description("Creates and returns an ADODB.Command, without executing it.")
Public Function CreateCommand(ByVal db As IDbConnection, ByVal commandType As ADODB.CommandTypeEnum, ByVal sql As String, ByRef args() As Variant) As ADODB.Command
Attribute CreateCommand.VB_Description = "Creates and returns an ADODB.Command, without executing it."
End Function

'@Description("Executes specified ADODB.Command, and returns a disconnected recordset containing the results.")
Public Function GetDisconnectedRecordset(ByVal cmd As ADODB.Command) As ADODB.Recordset
Attribute GetDisconnectedRecordset.VB_Description = "Executes specified ADODB.Command, and returns a disconnected recordset containing the results."
End Function

'@Description("Executes specified ADODB.Command, and returns the value of the first field of the first returned record.")
Public Function GetSingleValue(ByVal db As IDbConnection, ByVal sql As String, ByRef args() As Variant) As Variant
Attribute GetSingleValue.VB_Description = "Executes specified ADODB.Command, and returns the value of the first field of the first returned record."
End Function

'@Description("True if the number of provided arguments matches the expected number of parameters in the 'sql' command string.")
Public Function ValidateOrdinalArguments(ByVal sql As String, ByRef args() As Variant) As Boolean
Attribute ValidateOrdinalArguments.VB_Description = "True if the number of provided arguments matches the expected number of parameters in the 'sql' command string."
End Function
