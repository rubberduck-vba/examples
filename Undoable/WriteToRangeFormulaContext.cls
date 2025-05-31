VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WriteToRangeFormulaContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Encapsulates the model for a WriteToRangeFormulaCommand"
'@Folder("Undoable.Commands.WriteToRangeFormula")
'@ModuleDescription("Encapsulates the model for a WriteToRangeFormulaCommand")
Option Explicit
Implements ICommandContext

Private Type TContext
    Target As Excel.Range
    Formula As String
End Type
Private This As TContext

'@Description("The target Range")
Public Property Get Target() As Excel.Range
Attribute Target.VB_Description = "The target Range"
    Set Target = This.Target
End Property

Public Property Set Target(ByVal RHS As Excel.Range)
    Set This.Target = RHS
End Property

'@Description("The formula or value to be written to the target")
Public Property Get Formula() As String
Attribute Formula.VB_Description = "The formula or value to be written to the target"
    Formula = This.Formula
End Property

Public Property Let Formula(ByVal RHS As String)
    This.Formula = RHS
End Property

Private Function ICommandContext_IsValid() As Boolean
    If Not This.Target Is Nothing Then
        If This.Target.Areas.Count = 1 Then
            ICommandContext_IsValid = True
        End If
    End If
End Function
