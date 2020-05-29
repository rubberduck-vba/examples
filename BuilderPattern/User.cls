VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")

Option Explicit
Implements IUser

Private Type TUser
    Id As String
    UserName As String
    FirstName As String
    LastName As String
    Email As String
    EmailVerified As Boolean
    TwoFactorEnabled As Boolean
    PhoneNumber As String
    PhoneNumberVerified As Boolean
    AvatarUrl As String
End Type

Private this As TUser

Public Property Get Id() As String
    Id = this.Id
End Property

Public Property Let Id(ByVal value As String)
    this.Id = value
End Property

Public Property Get UserName() As String
    UserName = this.UserName
End Property

Public Property Let UserName(ByVal value As String)
    this.UserName = value
End Property

Public Property Get FirstName() As String
    FirstName = this.FirstName
End Property

Public Property Let FirstName(ByVal value As String)
    this.FirstName = value
End Property

Public Property Get LastName() As String
    LastName = this.LastName
End Property

Public Property Let LastName(ByVal value As String)
    this.LastName = value
End Property

Public Property Get Email() As String
    Email = this.Email
End Property

Public Property Let Email(ByVal value As String)
    this.Email = value
End Property

Public Property Get EmailVerified() As Boolean
    EmailVerified = this.EmailVerified
End Property

Public Property Let EmailVerified(ByVal value As Boolean)
    this.EmailVerified = value
End Property

Public Property Get TwoFactorEnabled() As Boolean
    TwoFactorEnabled = this.TwoFactorEnabled
End Property

Public Property Let TwoFactorEnabled(ByVal value As Boolean)
    this.TwoFactorEnabled = value
End Property

Public Property Get PhoneNumber() As String
    PhoneNumber = this.PhoneNumber
End Property

Public Property Let PhoneNumber(ByVal value As String)
    this.PhoneNumber = value
End Property

Public Property Get PhoneNumberVerified() As Boolean
    PhoneNumberVerified = this.PhoneNumberVerified
End Property

Public Property Let PhoneNumberVerified(ByVal value As Boolean)
    this.PhoneNumberVerified = value
End Property

Public Property Get AvatarUrl() As String
    AvatarUrl = this.AvatarUrl
End Property

Public Property Let AvatarUrl(ByVal value As String)
    this.AvatarUrl = value
End Property

Private Property Get IUser_Id() As String
    IUser_Id = this.Id
End Property

Private Property Get IUser_UserName() As String
    IUser_UserName = this.UserName
End Property

Private Property Get IUser_FirstName() As String
    IUser_FirstName = this.FirstName
End Property

Private Property Get IUser_LastName() As String
    IUser_LastName = this.LastName
End Property

Private Property Get IUser_Email() As String
    IUser_Email = this.Email
End Property

Private Property Get IUser_EmailVerified() As Boolean
    IUser_EmailVerified = this.EmailVerified
End Property

Private Property Get IUser_TwoFactorEnabled() As Boolean
    IUser_TwoFactorEnabled = this.TwoFactorEnabled
End Property

Private Property Get IUser_PhoneNumber() As String
    IUser_PhoneNumber = this.PhoneNumber
End Property

Private Property Get IUser_PhoneNumberVerified() As Boolean
    IUser_PhoneNumberVerified = this.PhoneNumberVerified
End Property

Private Property Get IUser_AvatarUrl() As String
    IUser_AvatarUrl = this.AvatarUrl
End Property

