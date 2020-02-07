VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AuthPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("Authentication")
'@ModuleDescription("Represents an object responsible for authenticating the current user.")
Option Explicit
Implements IAuthPresenter

Private Type TPresenter
    View As IDialogView
    AuthService As IAuthService
    IsAuthenticated As Boolean
End Type
Private this As TPresenter

Public Function Create(ByVal service As IAuthService, ByVal dialogView As IDialogView) As IAuthPresenter
    Dim result As AuthPresenter
    Set result = New AuthPresenter
    Set result.AuthService = service
    Set result.View = dialogView
    Set Create = result
End Function

Public Property Get AuthService() As IAuthService
    Set AuthService = this.AuthService
End Property

Public Property Set AuthService(ByVal value As IAuthService)
    Set this.AuthService = value
End Property

Public Property Get View() As IDialogView
    Set View = this.View
End Property

Public Property Set View(ByVal value As IDialogView)
    Set this.View = value
End Property

Private Function IAuthPresenter_Authenticate() As Boolean
    If Not this.View.ShowDialog Then Exit Function
    this.IsAuthenticated = this.AuthService.Authenticate(this.View.ViewModel)
    IAuthPresenter_Authenticate = this.IsAuthenticated
End Function

Private Property Get IAuthPresenter_IsAuthenticated() As Boolean
    IAuthPresenter_IsAuthenticated = this.IsAuthenticated
End Property

