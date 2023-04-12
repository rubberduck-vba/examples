VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertyChangeNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Handlers As VBA.Collection

Public Sub AddHandler(ByVal Handler As IHandlePropertyChanged)
    Handlers.Add Handler
End Sub

Public Sub Notify(ByVal Source As Object, ByVal Name As String)
    Dim Handler As IHandlePropertyChanged
    For Each Handler In Handlers
        Handler.OnPropertyChanged Source, Name
    Next
End Sub

Private Sub Class_Initialize()
    Set Handlers = New VBA.Collection
End Sub
