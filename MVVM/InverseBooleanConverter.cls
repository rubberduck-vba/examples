VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InverseBooleanConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A value converter that inverts a Boolean value."
'@Folder MVVM.Infrastructure.Bindings.Converters
'@ModuleDescription "A value converter that inverts a Boolean value."
'@PredeclaredId
'@Exposed
Option Explicit
Implements IValueConverter

Public Function Default() As IValueConverter
    GuardClauses.GuardNonDefaultInstance Me, InverseBooleanConverter
    Set Default = InverseBooleanConverter
End Function

Private Function IValueConverter_Convert(ByVal Value As Variant) As Variant
    IValueConverter_Convert = Not CBool(Value)
End Function

Private Function IValueConverter_ConvertBack(ByVal Value As Variant) As Variant
    IValueConverter_ConvertBack = Not CBool(Value)
End Function
