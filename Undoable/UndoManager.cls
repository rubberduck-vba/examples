VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UndoManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("Undoable")
Option Explicit
Private UndoStack As Collection
Private RedoStack As Collection

Public Sub Clear()
    Do While UndoStack.Count > 0
        UndoStack.Remove 1
    Loop
    Do While RedoStack.Count > 0
        RedoStack.Remove 1
    Loop
End Sub

Public Sub Push(ByVal Action As IUndoable)
    ThrowOnInvalidInstance
    UndoStack.Add Action
End Sub

Public Function PopUndoStack() As IUndoable
    ThrowOnInvalidInstance
    
    Dim Item As IUndoable
    Set Item = UndoStack.Item(UndoStack.Count)
    
    UndoStack.Remove UndoStack.Count
    RedoStack.Add Item
    
    Set PopUndoStack = Item
End Function

Public Function PopRedoStack() As IUndoable
    ThrowOnInvalidInstance
    
    Dim Item As IUndoable
    Set Item = RedoStack.Item(RedoStack.Count)
    
    RedoStack.Remove RedoStack.Count
    UndoStack.Add Item
    
    Set PopRedoStack = Item
End Function

Public Property Get CanUndo() As Boolean
    CanUndo = UndoStack.Count > 0
End Property

Public Property Get CanRedo() As Boolean
    CanRedo = RedoStack.Count > 0
End Property

Public Property Get UndoState() As Variant
    If Not CanUndo Then Exit Sub
    ReDim Items(1 To UndoStack.Count) As String
    Dim StackIndex As Long
    For StackIndex = 1 To UndoStack.Count
        Dim Item As IUndoable
        Set Item = UndoStack.Item(StackIndex)
        Items(StackIndex) = StackIndex & vbTab & Item.Description
    Next
    UndoState = Items
End Property

Public Property Get RedoState() As Variant
    If Not CanRedo Then Exit Property
    ReDim Items(1 To RedoStack.Count) As String
    Dim StackIndex As Long
    For StackIndex = 1 To RedoStack.Count
        Dim Item As IUndoable
        Set Item = RedoStack.Item(StackIndex)
        Items(StackIndex) = StackIndex & vbTab & Item.Description
    Next
    RedoState = Items
End Property

Private Sub ThrowOnInvalidInstance()
    If Not Me Is UndoManager Then Err.Raise 5, TypeName(Me), "Instance is invalid"
End Sub

Private Sub Class_Initialize()
    Set UndoStack = New Collection
    Set RedoStack = New Collection
End Sub

Private Sub Class_Terminate()
    Set UndoStack = Nothing
    Set RedoStack = Nothing
End Sub
