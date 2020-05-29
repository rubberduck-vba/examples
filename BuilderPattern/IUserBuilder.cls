VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IUserBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Incrementally builds a User instance."
'@Interface
'@ModuleDescription("Incrementally builds a User instance.")
Option Explicit

'@Description("Returns the current object.")
Public Function Build() As IUser
Attribute Build.VB_Description = "Returns the current object."
End Function

'@Description("Builds a user with a first and last name.")
Public Function WithName(ByVal FirstName As String, ByVal LastName As String) As IUserBuilder
Attribute WithName.VB_Description = "Builds a user with a first and last name."
End Function

'@Description("Builds a user with an email address.")
Public Function WithEmail(ByVal Email As String, Optional ByVal Verified As Boolean = False) As IUserBuilder
Attribute WithEmail.VB_Description = "Builds a user with an email address."
End Function

'@Description("Builds a user with SMS-based 2FA enabled.")
Public Function WithTwoFactorAuthentication(ByVal PhoneNumber As String, Optional ByVal Verified As Boolean = False) As IUserBuilder
Attribute WithTwoFactorAuthentication.VB_Description = "Builds a user with SMS-based 2FA enabled."
End Function

'@Description("Builds a user with an avatar at the specified URL.")
Public Function WithAvatar(ByVal Url As String) As IUserBuilder
Attribute WithAvatar.VB_Description = "Builds a user with an avatar at the specified URL."
End Function
