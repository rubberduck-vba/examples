VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Car"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("Examples.Model")
'@Exposed
Option Explicit

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

