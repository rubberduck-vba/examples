VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an object abstracting a database command."
'@Folder("SecureADODB.DbCommand.Abstract")
'@ModuleDescription("Represents an object abstracting a database command.")
'@Interface
'@Exposed
Option Explicit

'@Description("Returns an ADODB.Recordset that contains the results of the specified parameterized query.")
Public Function Execute(ByVal sql As String, ParamArray args()) As ADODB.Recordset
Attribute Execute.VB_Description = "Returns an ADODB.Recordset that contains the results of the specified parameterized query."
End Function

'@Description("Executes the specified parameterized SQL statement/command.")
Public Sub ExecuteNonQuery(ByVal sql As String, ParamArray args())
Attribute ExecuteNonQuery.VB_Description = "Executes the specified parameterized SQL statement/command."
End Sub

'@Description("Returns the value of the first field of the first record returned by the specified parameterized query.")
Public Function GetSingleValue(ByVal sql As String, ParamArray args()) As Variant
Attribute GetSingleValue.VB_Description = "Returns the value of the first field of the first record returned by the specified parameterized query."
End Function

'@Description("Returns an ADODB.Recordset that contains the results of the specified parameterized query, using specified ADODB.Parameters.")
Public Function ExecuteWithParameters(ByVal sql As String, ParamArray params()) As ADODB.Recordset
Attribute ExecuteWithParameters.VB_Description = "Returns an ADODB.Recordset that contains the results of the specified parameterized query, using specified ADODB.Parameters."
End Function
