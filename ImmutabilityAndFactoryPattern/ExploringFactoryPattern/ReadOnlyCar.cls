VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReadOnlyCar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("Examples.Model")
'@PredeclaredId
'@Exposed
Option Explicit
Implements ICar

Private Type TCar
    Make As Long
    Model As String
    Manufacturer As String
End Type
Private this As TCar

Public Property Get Make() As Long
    Make = this.Make
End Property

Friend Property Let Make(ByVal value As Long)
    this.Make = value
End Property

Public Property Get Model() As String
    Model = this.Model
End Property

Friend Property Let Model(ByVal value As String)
    this.Model = value
End Property

Public Property Get Manufacturer() As String
    Manufacturer = this.Manufacturer
End Property

Friend Property Let Manufacturer(ByVal value As String)
    this.Manufacturer = value
End Property

Private Property Get ICar_Make() As Long
    ICar_Make = this.Make
End Property

Private Property Get ICar_Manufacturer() As String
    ICar_Manufacturer = this.Manufacturer
End Property

Private Property Get ICar_Model() As String
    ICar_Model = this.Model
End Property
