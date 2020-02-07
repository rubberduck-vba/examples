VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetAuthService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Authentication")
'@ModuleDescription("A service responsible for authentication.")
Option Explicit
Implements IAuthService

Private Type TAuthService
    Algorithm As IHashAlgorithm
End Type

Private Type TUserAuthInfo
    Salt As String
    Hash As String
End Type

Private this As TAuthService

Public Function Create(ByVal hashAlgorithm As IHashAlgorithm) As IAuthService
    With New WorksheetAuthService
        Set .Algorithm = hashAlgorithm
        Set Create = .Self
    End With
End Function

Public Property Get Self() As IAuthService
    Set Self = Me
End Property

Public Property Get Algorithm() As IHashAlgorithm
    Set Algorithm = this.Algorithm
End Property

Public Property Set Algorithm(ByVal value As IHashAlgorithm)
    Set this.Algorithm = value
End Property

Private Function GetUserAuthInfo(ByVal user As String, ByRef outInfo As TUserAuthInfo) As Boolean
'gets the salt value & password hash for the specified user; returns false if user can't be retrieved.
    On Error GoTo CleanFail
    With PasswordsSheet.Table
        
        Dim nameColumnIndex As Long
        nameColumnIndex = .ListColumns("Name").Index
        
        Dim saltColumnIndex As Long
        saltColumnIndex = .ListColumns("Salt").Index
        
        Dim hashColumnIndex As Long
        hashColumnIndex = .ListColumns("PasswordHash").Index
        
        Dim userRowIndex As Long
        userRowIndex = Application.WorksheetFunction.Match(user, .ListColumns(nameColumnIndex).DataBodyRange, 0)
        
        outInfo.Salt = Application.WorksheetFunction.Index(.ListColumns(saltColumnIndex).DataBodyRange, userRowIndex)
        outInfo.Hash = Application.WorksheetFunction.Index(.ListColumns(hashColumnIndex).DataBodyRange, userRowIndex)
    
    End With
    GetUserAuthInfo = True
CleanExit:
    Exit Function
CleanFail:
    Debug.Print Err.Description
    Debug.Print "Unable to retrieve authentication info for user '" & user & "'."
    outInfo.Salt = vbNullString
    outInfo.Hash = vbNullString
    GetUserAuthInfo = False
    Resume CleanExit
End Function

Private Function IAuthService_Authenticate(ByVal model As UserAuthModel) As Boolean
    
    Dim info As TUserAuthInfo
    If Not GetUserAuthInfo(model.Name, outInfo:=info) Then Exit Function
    
    Dim pwdHash As String
    pwdHash = this.Algorithm.ComputeHash(info.Salt & model.Password)
    
    IAuthService_Authenticate = (pwdHash = info.Hash)
    
End Function
