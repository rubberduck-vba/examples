VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AuthDialogView 
   Caption         =   "Authentication"
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "AuthDialogView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AuthDialogView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Authentication")
Option Explicit
Implements IDialogView
Private Type TAuthDialog
    UserAuthModel As UserAuthModel
    IsCancelled As Boolean
End Type
Private this As TAuthDialog

Public Function Create(ByVal model As UserAuthModel) As IDialogView
    If model Is Nothing Then Err.Raise 5, TypeName(Me), "Model cannot be a null reference"
    Dim result As AuthDialogView
    Set result = New AuthDialogView
    Set result.UserAuthModel = model
    Set Create = result
End Function

Public Property Get UserAuthModel() As UserAuthModel
    Set UserAuthModel = this.UserAuthModel
End Property

Public Property Set UserAuthModel(ByVal value As UserAuthModel)
    Set this.UserAuthModel = value
End Property

Private Sub OnCancel()
    this.IsCancelled = True
    Me.Hide
End Sub

Private Sub Validate()
    OkButton.Enabled = this.UserAuthModel.IsValid
End Sub

Private Sub CancelButton_Click()
    OnCancel
End Sub

Private Sub OkButton_Click()
    Me.Hide
End Sub

Private Sub NameBox_Change()
    this.UserAuthModel.Name = NameBox.Text
    Validate
End Sub

Private Sub PasswordBox_Change()
    this.UserAuthModel.Password = PasswordBox.Text
    Validate
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Function IDialogView_ShowDialog() As Boolean
    Me.Show vbModal
    IDialogView_ShowDialog = Not this.IsCancelled
End Function

Private Property Get IDialogView_ViewModel() As Object
    Set IDialogView_ViewModel = this.UserAuthModel
End Property

