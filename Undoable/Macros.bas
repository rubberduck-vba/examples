Attribute VB_Name = "Macros"
'@Folder("UndoCmd")
Option Explicit

'@EntryPoint
Public Sub DoSomething()
    UndoManager.Clear
    CommandManager.WriteToFormula Sheet1.Range("A1"), "Hello"
    CommandManager.WriteToFormula Sheet1.Range("B1"), "World!"
    CommandManager.WriteToFormula Sheet1.Range("C1:C10"), "=RANDBETWEEN(0, 255)"
    CommandManager.WriteToFormula Sheet1.Range("D1:D10"), "=SUM($C$1:$C1)"
    CommandManager.SetNumberFormat Sheet1.Range("D1:D10"), "$#,##0.00"
End Sub

'@EntryPoint
Public Sub UndoAll()
    CommandManager.UndoAll
End Sub

'@EntryPoint
Public Sub UndoAction()
    CommandManager.UndoAction
End Sub

'@EntryPoint
Public Sub RedoAction()
    CommandManager.RedoAction
End Sub

'@EntryPoint
Public Sub ShowState()
    Dim UndoState As String
    If UndoManager.CanUndo Then
        UndoState = Join(UndoManager.UndoState, vbNewLine)
    Else
        UndoState = "(undo stack is empty)"
    End If
    
    Dim RedoState As String
    If UndoManager.CanRedo Then
        RedoState = Join(UndoManager.RedoState, vbNewLine)
    Else
        RedoState = "(redo stack is empty)"
    End If
    
    MsgBox "Undo state:" & vbNewLine & UndoState & vbNewLine & vbNewLine & _
           "Redo state:" & vbNewLine & RedoState
End Sub

