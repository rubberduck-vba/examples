VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("Undoable.Commands")
Option Explicit

Public Sub WriteToFormula(ByVal Target As Range, ByVal Formula As String)
    Dim Command As ICommand
    Set Command = New WriteToRangeFormulaCommand
    
    Dim Context As WriteToRangeFormulaContext
    Set Context = New WriteToRangeFormulaContext
    
    Set Context.Target = Target
    Context.Formula = Formula
    
    RunCommand Command, Context
End Sub

Public Sub SetNumberFormat(ByVal Target As Range, ByVal FormatString As String)
    Dim Command As ICommand
    Set Command = New SetNumberFormatCommand
    
    Dim Context As SetNumberFormatContext
    Set Context = New SetNumberFormatContext
    
    Set Context.Target = Target
    Context.FormatString = FormatString
    
    RunCommand Command, Context
End Sub

Public Sub UndoAction()
    If UndoManager.CanUndo Then UndoManager.PopUndoStack.Undo
End Sub

Public Sub UndoAll()
    Do While UndoManager.CanUndo
        UndoManager.PopUndoStack.Undo
    Loop
End Sub

Public Sub RedoAction()
    If UndoManager.CanRedo Then UndoManager.PopRedoStack.Redo
End Sub

Public Sub RedoAll()
    Do While UndoManager.CanRedo
        UndoManager.PopRedoStack.Redo
    Loop
End Sub

Public Property Get CanUndo() As Boolean
    CanUndo = UndoManager.CanUndo
End Property

Public Property Get CanRedo() As Boolean
    CanRedo = UndoManager.CanRedo
End Property

Private Sub RunCommand(ByVal Command As ICommand, ByVal Context As ICommandContext)
    If Command.CanExecute(Context) Then
        Command.Execute Context
        StackUndoable Command
    Else
        Debug.Print "Command cannot be executed in this context."
    End If
End Sub

Private Sub ThrowOnInvalidInstance()
    If Not Me Is CommandManager Then Err.Raise 5, TypeName(Me), "Instance is invalid"
End Sub

Private Sub StackUndoable(ByVal Command As Object)
    If TypeOf Command Is IUndoable Then
        Dim Undoable As IUndoable
        Set Undoable = Command
        UndoManager.Push Undoable
    End If
End Sub

