VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WriteToRangeFormulaCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "An undoable command that writes to the Formula2 property of a provided Range target"
'@Folder("Undoable.Commands.WriteToRangeFormula")
'@ModuleDescription("An undoable command that writes to the Formula2 property of a provided Range target")
Option Explicit
Implements ICommand
Implements IUndoable

Private Type TState
    InitialFormulas As Variant
    Context As WriteToRangeFormulaContext
    DidRun As Boolean
    DidUndo As Boolean
End Type

Private This As TState

Private Function ICommand_CanExecute(ByVal Context As Object) As Boolean
    ICommand_CanExecute = CanExecuteInternal(Context)
End Function

Private Sub ICommand_Execute(ByVal Context As Object)
    ExecuteInternal Context
End Sub

Private Property Get IUndoable_Description() As String
    IUndoable_Description = GetDescriptionInternal
End Property

Private Sub IUndoable_Redo()
    RedoInternal
End Sub

Private Sub IUndoable_Undo()
    UndoInternal
End Sub

Private Function GetDescriptionInternal() As String
    Dim FormulaText As String
    If Len(This.Context.Formula) > 20 Then
        FormulaText = "formula"
    Else
        FormulaText = "'" & This.Context.Formula & "'"
    End If
    GetDescriptionInternal = "Write " & FormulaText & " to " & This.Context.Target.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
End Function

Private Function CanExecuteInternal(ByVal Context As Object) As Boolean
    On Error GoTo OnInvalidContext
    
    GuardInvalidContext Context
    CanExecuteInternal = Not This.DidRun
    
    Exit Function
OnInvalidContext:
    CanExecuteInternal = False
End Function

Private Sub ExecuteInternal(ByVal Context As WriteToRangeFormulaContext)
    
    GuardInvalidContext Context
    SetUndoState Context
    
    Debug.Print "> Executing action: " & GetDescriptionInternal
    
    Context.Target.Formula2 = Context.Formula
    This.DidRun = True
    
End Sub

Private Sub GuardInvalidContext(ByVal Context As Object)
    If Not TypeOf Context Is ICommandContext Then Err.Raise 5, TypeName(Me), "An invalid context type was provided."
    Dim SafeContext As ICommandContext
    Set SafeContext = Context
    If Not SafeContext.IsValid And Not TypeOf Context Is WriteToRangeFormulaContext Then Err.Raise 5, TypeName(Me), "An invalid context was provided."
End Sub

Private Sub SetUndoState(ByVal Context As WriteToRangeFormulaContext)
    Set This.Context = Context
    This.InitialFormulas = Context.Target.Formula2
End Sub

Private Sub UndoInternal()
    If Not This.DidRun Then Err.Raise 5, TypeName(Me), "Cannot undo what has not been done."
    If This.DidUndo Then Err.Raise 5, TypeName(Me), "Operation was already undone."
    
    Debug.Print "> Undoing action: " & GetDescriptionInternal
    
    This.Context.Target.Formula2 = This.InitialFormulas
    This.DidUndo = True
End Sub

Private Sub RedoInternal()
    If Not This.DidUndo Then Err.Raise 5, TypeName(Me), "Cannot redo what was never undone."
    ExecuteInternal This.Context
    This.DidUndo = False
End Sub

