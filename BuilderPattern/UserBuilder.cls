VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Builds a User object."
'@PredeclaredId
'@ModuleDescription("Builds a User object.")
Option Explicit
Implements IUserBuilder
Private internal As User

'@Description("Creates a new UserBuilder instance.")
Public Function Create(ByVal Id As String, ByVal UserName As String) As IUserBuilder
Attribute Create.VB_Description = "Creates a new UserBuilder instance."
    Dim result As UserBuilder
    Set result = New UserBuilder
    
    '@Ignore UserMeaningfulName FIXME
    Dim obj As User
    Set obj = New User
    obj.Id = Id
    obj.UserName = UserName
    
    Set result.User = internal
    Set Create = result
End Function

'@Ignore WriteOnlyProperty
'@Description("For property injection of the internal IUser object; only the Create method should be invoking this member.")
Friend Property Set User(ByVal value As IUser)
Attribute User.VB_Description = "For property injection of the internal IUser object; only the Create method should be invoking this member."
    If Me Is UserBuilder Then Err.Raise 5, TypeName(Me), "Member call is illegal from default instance."
    If value Is Nothing Then Err.Raise 5, TypeName(Me), "'value' argument cannot be a null reference."
    Set internal = value
End Property

Private Function IUserBuilder_Build() As IUser
    If internal Is Nothing Then Err.Raise 91, TypeName(Me), "Builder initialization error: use UserBuilder.Create to create a UserBuilder."
    Set IUserBuilder_Build = internal
End Function

Private Function IUserBuilder_WithName(ByVal FirstName As String, ByVal LastName As String) As IUserBuilder
    internal.FirstName = FirstName
    internal.LastName = LastName
    Set IUserBuilder_WithName = Me
End Function

Private Function IUserBuilder_WithEmail(ByVal Email As String, Optional ByVal Verified As Boolean = False) As IUserBuilder
    internal.Email = Email
    internal.EmailVerified = Verified
    Set IUserBuilder_WithEmail = Me
End Function

Private Function IUserBuilder_WithTwoFactorAuthentication(ByVal PhoneNumber As String, Optional ByVal Verified As Boolean = False) As IUserBuilder
    internal.TwoFactorEnabled = True
    internal.PhoneNumber = PhoneNumber
    internal.PhoneNumberVerified = Verified
    Set IUserBuilder_WithTwoFactorAuthentication = Me
End Function

Private Function IUserBuilder_WithAvatar(ByVal Url As String) As IUserBuilder
    internal.AvatarUrl = Url
    Set IUserBuilder_WithAvatar = Me
End Function
