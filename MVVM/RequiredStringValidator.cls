VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RequiredStringValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder MVVM.Example
Option Explicit
Implements IValueValidator

Private Function IValueValidator_IsValid(ByVal Value As Variant, ByVal Source As Object, ByVal Target As Object) As Boolean
    IValueValidator_IsValid = Len(Trim$(Value)) > 0
End Function

Private Property Get IValueValidator_Message() As String
    IValueValidator_Message = "Value cannot be empty."
End Property
