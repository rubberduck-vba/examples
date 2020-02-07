VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserAuthModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("Authentication")
Option Explicit

Private Type TAuthModel
    Name As String
    Password As String
    IsValid As Boolean
End Type
Private this As TAuthModel

Public Property Get Name() As String
    Name = this.Name
End Property

Public Property Let Name(ByVal value As String)
    this.Name = value
    Validate
End Property

Public Property Get Password() As String
    Password = this.Password
End Property

Public Property Let Password(ByVal value As String)
    this.Password = value
    Validate
End Property

Public Property Get IsValid() As Boolean
    IsValid = this.IsValid
End Property

Private Sub Validate()
    this.IsValid = Len(this.Name) > 0 And Len(this.Password) > 0
End Sub
