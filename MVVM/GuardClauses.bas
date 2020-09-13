Attribute VB_Name = "GuardClauses"
Attribute VB_Description = "Global procedures for throwing custom run-time errors in guard clauses."
'@Folder MVVM.Errors
'@ModuleDescription("Global procedures for throwing custom run-time errors in guard clauses.")
Option Explicit
Option Private Module

'@Description("Raises a run-time error if the specified Boolean expression is True.")
Public Sub GuardExpression(ByVal Throw As Boolean, _
Optional ByVal Source As String = vbNullString, _
Optional ByVal Message As String = "Invalid procedure call or argument.")
Attribute GuardExpression.VB_Description = "Raises a run-time error if the specified Boolean expression is True."
    If Throw Then VBA.Information.Err.Raise CustomError, Source, Message
End Sub

'@Description("Raises a run-time error if the specified instance isn't the default instance.")
Public Sub GuardNonDefaultInstance(ByVal Instance As Object, ByVal defaultInstance As Object, _
Optional ByVal Source As String = vbNullString, _
Optional ByVal Message As String = "Method should be invoked from the default/predeclared instance of this class.")
Attribute GuardNonDefaultInstance.VB_Description = "Raises a run-time error if the specified instance isn't the default instance."
    Debug.Assert TypeName(Instance) = TypeName(defaultInstance)
    GuardExpression Not Instance Is defaultInstance, Source, Message
End Sub

'@Description("Raises a run-time error if the specified object reference is already set.")
Public Sub GuardDoubleInitialization(ByVal Instance As Object, _
Optional ByVal Source As String = vbNullString, _
Optional ByVal Message As String = "Object is already initialized.")
Attribute GuardDoubleInitialization.VB_Description = "Raises a run-time error if the specified object reference is already set."
    GuardExpression Not Instance Is Nothing, Source, Message
End Sub

'@Description("Raises a run-time error if the specified object reference is Nothing.")
Public Sub GuardNullReference(ByVal Instance As Object, _
Optional ByVal Source As String = vbNullString, _
Optional ByVal Message As String = "Object reference cannot be Nothing.")
Attribute GuardNullReference.VB_Description = "Raises a run-time error if the specified object reference is Nothing."
    GuardExpression Instance Is Nothing, Source, Message
End Sub

'@Description("Raises a run-time error if the specified string is empty.")
Public Sub GuardEmptyString(ByVal Value As String, _
Optional ByVal Source As String = vbNullString, _
Optional ByVal Message As String = "String cannot be empty.")
Attribute GuardEmptyString.VB_Description = "Raises a run-time error if the specified string is empty."
    GuardExpression Value = vbNullString, Source, Message
End Sub
