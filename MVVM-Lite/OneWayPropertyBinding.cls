VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OneWayPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IHandlePropertyChanged

Private Type TBinding
    InvertBoolean As Boolean
    Source As Object
    SourceProperty As String
    Target As Object
    TargetProperty As String
End Type

Private This As TBinding

Public Sub Initialize(ByVal Target As MSForms.Control, ByVal TargetProperty As String, ByVal Source As Object, ByVal SourceProperty As String, Optional ByVal InvertBoolean As Boolean = False)
    Set This.Source = Source
    This.SourceProperty = SourceProperty
    Set This.Target = Target
    This.TargetProperty = TargetProperty
    This.InvertBoolean = InvertBoolean
    If TypeOf Source Is INotifyPropertyChanged Then RegisterPropertyChanges Source
    IHandlePropertyChanged_OnPropertyChanged Source, SourceProperty
End Sub

Private Sub RegisterPropertyChanges(ByVal Source As INotifyPropertyChanged)
    Source.RegisterHandler Me
End Sub

Private Sub IHandlePropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal Name As String)
    If Source Is This.Source And Name = This.SourceProperty Then
        Dim Value As Variant
        Value = VBA.Interaction.CallByName(This.Source, This.SourceProperty, VbGet)
        If VarType(Value) = vbBoolean And This.InvertBoolean Then
            VBA.Interaction.CallByName This.Target, This.TargetProperty, VbLet, Not CBool(Value)
        Else
            VBA.Interaction.CallByName This.Target, This.TargetProperty, VbLet, Value
        End If
    End If
End Sub

