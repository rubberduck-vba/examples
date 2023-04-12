VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComboBoxValueBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IHandlePropertyChanged
Private WithEvents UI As MSForms.ComboBox
Attribute UI.VB_VarHelpID = -1

Private Type TBinding
    Source As Object
    SourceProperty As String
End Type

Private This As TBinding

Public Sub Initialize(ByVal Control As MSForms.ComboBox, ByVal Source As Object, ByVal SourceProperty As String)
    Set UI = Control
    Set This.Source = Source
    This.SourceProperty = SourceProperty
    If TypeOf Source Is INotifyPropertyChanged Then RegisterPropertyChanges Source
End Sub

Private Sub RegisterPropertyChanges(ByVal Source As INotifyPropertyChanged)
    Source.RegisterHandler Me
End Sub

Private Sub IHandlePropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal Name As String)
    If Source Is This.Source And Name = This.SourceProperty Then
        On Error Resume Next
        UI.Text = VBA.Interaction.CallByName(This.Source, This.SourceProperty, VbGet)
        On Error GoTo 0
    End If
End Sub

Private Sub UI_Change()
    VBA.Interaction.CallByName This.Source, This.SourceProperty, VbLet, UI.Value
End Sub

