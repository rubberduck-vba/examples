VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SetNumberFormatContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Undoable.Commands.SetNumberFormat")
'@ModuleDescription("Encapsulates the model for a SetNumberFormatCommand")
Option Explicit
Implements ICommandContext

Private Type TContext
    Target As Excel.Range
    FormatString As String
End Type
Private This As TContext

'@Description("The target Range")
Public Property Get Target() As Excel.Range
    Set Target = This.Target
End Property

Public Property Set Target(ByVal RHS As Excel.Range)
    Set This.Target = RHS
End Property

'@Description("The format string applied to the target")
Public Property Get FormatString() As String
    FormatString = This.FormatString
End Property

Public Property Let FormatString(ByVal RHS As String)
    This.FormatString = RHS
End Property

Private Function ICommandContext_IsValid() As Boolean
    If Not This.Target Is Nothing Then
        If This.Target.Areas.Count = 1 Then
            ICommandContext_IsValid = True
        End If
    End If
End Function

