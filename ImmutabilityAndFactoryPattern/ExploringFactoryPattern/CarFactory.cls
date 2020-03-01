VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CarFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("Examples.ReadOnlyInReferencingProject.Factory")
'@PredeclaredId
'@Exposed
Option Explicit

Public Function Create(ByVal carMake As Long, ByVal carModel As String, ByVal carManufacturer As String) As Car
    Dim result As Car
    Set result = New Car
    result.Make = carMake
    result.Model = carModel
    result.Manufacturer = carManufacturer
    Set Create = result
End Function
