VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ManufacturerCarFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("Examples.ReadOnlyEverywhere.AbstractFactory")
'@PredeclaredId
Option Explicit
Implements ISimplerCarFactory

Private Type TFactory
    Manufacturer As String
End Type

Private this As TFactory

Public Function Create(ByVal carManufacturer As String) As ISimplerCarFactory
    Dim result As ManufacturerCarFactory
    Set result = New ManufacturerCarFactory
    result.Manufacturer = carManufacturer
    Set Create = result
End Function

Public Property Get Manufacturer() As String
    Manufacturer = this.Manufacturer
End Property

Friend Property Let Manufacturer(ByVal value As String)
    this.Manufacturer = value
End Property

Private Function ISimplerCarFactory_Create(ByVal carModel As String) As ICar
    Dim result As ReadOnlyCar
    Set result = New ReadOnlyCar
    result.Manufacturer = this.Manufacturer
    result.Make = DateTime.Year(DateTime.Date)
    result.Model = carModel
    Set ISimplerCarFactory_Create = result
End Function
