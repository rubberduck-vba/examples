VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OrderLineItemModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements INotifyPropertyChanged

Private Notification As New PropertyChangeNotification

Private Type TLineItem
    SKU As String
    Quantity As Long
    Price As Currency
End Type

Private This As TLineItem

Public Property Get SKU() As String
    SKU = This.SKU
End Property

Public Property Let SKU(ByVal Value As String)
    This.SKU = Value
    OnPropertyChanged "SKU"
End Property

Public Property Get Quantity() As Long
    Quantity = This.Quantity
End Property

Public Property Let Quantity(ByVal Value As Long)
    This.Quantity = Value
    OnPropertyChanged "Quantity"
End Property

Public Property Get Price() As Currency
    Price = This.Price
End Property

Public Property Let Price(ByVal Value As Currency)
    This.Price = Value
    OnPropertyChanged "Price"
End Property

Private Sub OnPropertyChanged(ByVal Name As String)
    INotifyPropertyChanged_OnPropertyChanged Me, Name
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal Name As String)
    Notification.Notify Source, Name
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    Notification.AddHandler Handler
End Sub

