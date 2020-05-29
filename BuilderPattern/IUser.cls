VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Folder("VBAProject")
'@Interface

Public Property Get Id() As String
End Property

Public Property Get UserName() As String
End Property

Public Property Get FirstName() As String
End Property

Public Property Get LastName() As String
End Property

Public Property Get Email() As String
End Property

Public Property Get EmailVerified() As Boolean
End Property

Public Property Get TwoFactorEnabled() As Boolean
End Property

Public Property Get PhoneNumber() As String
End Property

Public Property Get PhoneNumberVerified() As Boolean
End Property

Public Property Get AvatarUrl() As String
End Property

