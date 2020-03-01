VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AbstractCarFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("Examples.ReadOnlyEverywhere.AbstractFactory")
'@PredeclaredId
'@Exposed
Option Explicit
Implements ICarFactory

Private Function ICarFactory_Create(ByVal carMake As Long, ByVal carModel As String, ByVal carManufacturer As String) As ICar
    Dim result As ICar
    Set result = New ReadOnlyCar '<~ we could use a strategy pattern here and output a different implementation depending on parameters
    result.Make = carMake
    result.Model = carModel
    result.Manufacturer = carManufacturer
    Set ICarFactory_Create = result
End Function

