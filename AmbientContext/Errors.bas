Attribute VB_Name = "Errors"
Attribute VB_Description = "Global procedures for throwing common errors."
'@Folder("AmbientContext")
'@ModuleDescription("Global procedures for throwing common errors.")
Option Explicit
Option Private Module

Public Const CustomError As Long = vbObjectError Or 32

'@Description("Re-raises the current error, if there is one.")
Public Sub RethrowOnError()
Attribute RethrowOnError.VB_Description = "Re-raises the current error, if there is one."
    With VBA.Information.Err
        If .Number <> 0 Then
            Debug.Print "Error " & .Number, .Description
            .Raise .Number
        End If
    End With
End Sub

'@Description("Raises a run-time error if the specified Boolean expression is True.")
Public Sub GuardExpression(ByVal throw As Boolean, _
Optional ByVal Source As String = "Errors", _
Optional ByVal message As String = "Invalid procedure call or argument.")
Attribute GuardExpression.VB_Description = "Raises a run-time error if the specified Boolean expression is True."
    If throw Then VBA.Information.Err.Raise CustomError, Source, message
End Sub

'@Description("Raises a run-time error if the specified instance isn't the default instance.")
Public Sub GuardNonDefaultInstance(ByVal Instance As Object, ByVal defaultInstance As Object, _
Optional ByVal Source As String = "Errors", _
Optional ByVal message As String = "Method should be invoked from the default/predeclared instance of this class.")
Attribute GuardNonDefaultInstance.VB_Description = "Raises a run-time error if the specified instance isn't the default instance."
    Debug.Assert TypeName(Instance) = TypeName(defaultInstance)
    GuardExpression Not Instance Is defaultInstance, Source, message
End Sub

'@Description("Raises a run-time error if the specified object reference is already set.")
Public Sub GuardDoubleInitialization(ByVal Instance As Object, _
Optional ByVal Source As String = "Errors", _
Optional ByVal message As String = "Object is already initialized.")
Attribute GuardDoubleInitialization.VB_Description = "Raises a run-time error if the specified object reference is already set."
    GuardExpression Not Instance Is Nothing, Source, message
End Sub

'@Description("Raises a run-time error if the specified object reference is Nothing.")
Public Sub GuardNullReference(ByVal Instance As Object, _
Optional ByVal Source As String = "Errors", _
Optional ByVal message As String = "Object reference cannot be Nothing.")
Attribute GuardNullReference.VB_Description = "Raises a run-time error if the specified object reference is Nothing."
    GuardExpression Instance Is Nothing, Source, message
End Sub

'@Description("Raises a run-time error if the specified string is empty.")
Public Sub GuardEmptyString(ByVal Value As String, _
Optional ByVal Source As String = "Errors", _
Optional ByVal message As String = "String cannot be empty.")
Attribute GuardEmptyString.VB_Description = "Raises a run-time error if the specified string is empty."
    GuardExpression Value = vbNullString, Source, message
End Sub

