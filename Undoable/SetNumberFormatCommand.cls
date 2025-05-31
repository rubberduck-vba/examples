VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SetNumberFormatCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Undoable.Commands.SetNumberFormat")
'@ModuleDescription("An undoable command that sets the number format of a provided Range target")
Option Explicit
Implements ICommand
Implements IUndoable

Private Type TState
    InitialFormatString As String
    Context As SetNumberFormatContext
    DidRun As Boolean
    DidUndo As Boolean
End Type

Private This As TState

Private Function GetDescriptionInternal() As String
    GetDescriptionInternal = "Set number format for " & This.Context.Target.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False)
End Function

Private Function CanExecuteInternal(ByVal Context As Object) As Boolean
    On Error GoTo OnInvalidContext
    
    GuardInvalidContext Context
    CanExecuteInternal = Not This.DidRun
    
    Exit Function
OnInvalidContext:
    CanExecuteInternal = False
End Function

Private Sub ExecuteInternal(ByVal Context As SetNumberFormatContext)
    
    GuardInvalidContext Context
    SetUndoState Context
    
    Debug.Print "> Executing action: " & GetDescriptionInternal
    
    Context.Target.NumberFormat = Context.FormatString
    This.DidRun = True
    
End Sub

Private Sub GuardInvalidContext(ByVal Context As Object)
    If Not TypeOf Context Is ICommandContext Then Err.Raise 5, TypeName(Me), "An invalid context type was provided."
    Dim SafeContext As ICommandContext
    Set SafeContext = Context
    If Not SafeContext.IsValid And Not TypeOf Context Is SetNumberFormatContext Then Err.Raise 5, TypeName(Me), "An invalid context was provided."
End Sub

Private Sub SetUndoState(ByVal Context As SetNumberFormatContext)
    Set This.Context = Context
    This.InitialFormatString = Context.Target.NumberFormat
End Sub

Private Sub UndoInternal()
    If Not This.DidRun Then Err.Raise 5, TypeName(Me), "Cannot undo an action that has not run."
    If This.DidUndo Then Err.Raise 5, TypeName(Me), "Operation was already undone."
    
    Debug.Print "> Undoing action: " & GetDescriptionInternal
    
    This.Context.Target.NumberFormat = This.InitialFormatString
    This.DidUndo = True
End Sub

Private Sub RedoInternal()
    If Not This.DidUndo Then Err.Raise 5, TypeName(Me), "Cannot redo an action that was not previously undone."
    ExecuteInternal This.Context
    This.DidUndo = False
End Sub

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

