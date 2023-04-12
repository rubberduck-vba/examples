VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OrderHeaderModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements INotifyPropertyChanged

Private Notification As New PropertyChangeNotification

Private Type TAddress
    Name As String
    Line1 As String
    Line2 As String
    Line3 As String
End Type

Private Type TCustomerOrder
    OrderNumber As Long
    OrderDate As Date
    BillTo As TAddress
    ShipTo As TAddress
    ShipToBillingAddress As Boolean
    EmailAddress As String
    Phone As String
    IsContributor As Boolean
    LineItems As VBA.Collection
    
    AddingNewItem As Boolean
    NewLineItem As OrderLineItemModel
End Type

Private This As TCustomerOrder

Public Property Get OrderNumber() As Long
    OrderNumber = This.OrderNumber
End Property

Public Property Let OrderNumber(ByVal Value As Long)
    If This.OrderDate <> Value Then
        This.OrderNumber = Value
        OnPropertyChanged "OrderNumber"
    End If
End Property

Public Property Get OrderDate() As Date
    OrderDate = This.OrderDate
End Property

Public Property Let OrderDate(ByVal Value As Date)
    If This.OrderDate <> Value Then
        This.OrderDate = Value
        OnPropertyChanged "OrderDate"
    End If
End Property

Public Property Get EmailAddress() As String
    EmailAddress = This.EmailAddress
End Property

Public Property Let EmailAddress(ByVal Value As String)
    If This.EmailAddress <> Value Then
        This.EmailAddress = Value
        OnPropertyChanged "EmailAddress"
    End If
End Property

Public Property Get IsContributor() As Boolean
    IsContributor = This.IsContributor
End Property

Public Property Let IsContributor(ByVal Value As Boolean)
    If This.IsContributor <> Value Then
        This.IsContributor = Value
        OnPropertyChanged "IsContributor"
    End If
End Property

Public Property Get ShipToBillingAddress() As Boolean
    ShipToBillingAddress = This.ShipToBillingAddress
End Property

Public Property Let ShipToBillingAddress(ByVal Value As Boolean)
    If This.ShipToBillingAddress <> Value Then
        This.ShipToBillingAddress = Value
        OnPropertyChanged "ShipToBillingAddress"
        
        If This.ShipToBillingAddress Then
            ShipToName = This.BillTo.Name
            ShipToLine1 = This.BillTo.Line1
            ShipToLine2 = This.BillTo.Line2
            ShipToLine3 = This.BillTo.Line3
        End If
    End If
End Property

Public Property Get BillToName() As String
    BillToName = This.BillTo.Name
End Property

Public Property Let BillToName(ByVal Value As String)
    If This.BillTo.Name <> Value Then
        This.BillTo.Name = Value
        OnPropertyChanged "BillToName"
        If This.ShipToBillingAddress Then
            ShipToName = Value
        End If
    End If
End Property

Public Property Get BillToLine1() As String
    BillToLine1 = This.BillTo.Line1
End Property

Public Property Let BillToLine1(ByVal Value As String)
    If This.BillTo.Line1 <> Value Then
        This.BillTo.Line1 = Value
        OnPropertyChanged "BillToLine1"
        
        If This.ShipToBillingAddress Then
            ShipToLine1 = Value
        End If
        End If
End Property

Public Property Get BillToLine2() As String
    BillToLine2 = This.BillTo.Line2
End Property

Public Property Let BillToLine2(ByVal Value As String)
    If This.BillTo.Line2 <> Value Then
        This.BillTo.Line2 = Value
        OnPropertyChanged "BillToLine2"
        
        If This.ShipToBillingAddress Then
            ShipToLine2 = Value
        End If
    End If
End Property

Public Property Get BillToLine3() As String
    BillToLine3 = This.BillTo.Line3
End Property

Public Property Let BillToLine3(ByVal Value As String)
    If This.BillTo.Line3 <> Value Then
        This.BillTo.Line3 = Value
        OnPropertyChanged "BillToLine3"
        
        If This.ShipToBillingAddress Then
            ShipToLine3 = Value
        End If
    End If
End Property

Public Property Get ShipToName() As String
    ShipToName = This.ShipTo.Name
End Property

Public Property Let ShipToName(ByVal Value As String)
    If This.ShipTo.Name <> Value Then
        This.ShipTo.Name = Value
        OnPropertyChanged "ShipToName"
    End If
End Property

Public Property Get ShipToLine1() As String
    ShipToLine1 = This.ShipTo.Line1
End Property

Public Property Let ShipToLine1(ByVal Value As String)
    If This.ShipTo.Line1 <> Value Then
        This.ShipTo.Line1 = Value
        OnPropertyChanged "ShipToLine1"
    End If
End Property

Public Property Get ShipToLine2() As String
    ShipToLine2 = This.ShipTo.Line2
End Property

Public Property Let ShipToLine2(ByVal Value As String)
    If This.ShipTo.Line2 <> Value Then
        This.ShipTo.Line2 = Value
        OnPropertyChanged "ShipToLine2"
    End If
End Property

Public Property Get ShipToLine3() As String
    ShipToLine3 = This.ShipTo.Line3
End Property

Public Property Let ShipToLine3(ByVal Value As String)
    If This.ShipTo.Line3 <> Value Then
        This.ShipTo.Line3 = Value
        OnPropertyChanged "ShipToLine3"
    End If
End Property

Public Property Get ShipToPhone() As String
    ShipToPhone = This.Phone
End Property

Public Property Let ShipToPhone(ByVal Value As String)
    If This.Phone <> Value Then
        This.Phone = Value
        OnPropertyChanged "ShipToPhone"
    End If
End Property

Public Property Get AddingLineItem() As Boolean
    AddingLineItem = This.AddingNewItem
End Property

Public Property Let AddingLineItem(ByVal Value As Boolean)
    If This.AddingNewItem <> Value Then
        This.AddingNewItem = Value
        OnPropertyChanged "AddingLineItem"
    End If
End Property

Public Property Get LineItems() As Variant
    If This.LineItems.Count > 0 Then
        
        ReDim Items(1 To This.LineItems.Count, 1 To 3) As Variant
        
        Dim Current As OrderLineItemModel, Index As Long
        For Each Current In This.LineItems
            Index = Index + 1
            Items(Index, 1) = Current.SKU
            Items(Index, 2) = Current.Quantity
            Items(Index, 3) = Current.Price
        Next
        LineItems = Items
    Else
        LineItems = Array()
    End If
End Property

Public Property Get LineItemObjects()
    Set LineItemObjects = This.LineItems
End Property

Public Property Get NewLineItem() As OrderLineItemModel
    Set NewLineItem = This.NewLineItem
End Property

Public Sub AddLineItem(ByVal Item As OrderLineItemModel)
    This.LineItems.Add Item
    OnPropertyChanged "LineItems"
End Sub

Public Sub RemoveLineItem(ByVal Index As Long)
    On Error Resume Next
    This.LineItems.Remove Index
    On Error GoTo 0
    OnPropertyChanged "LineItems"
End Sub

Public Sub ClearLineItems()
    Set This.LineItems = New VBA.Collection
    OnPropertyChanged "LineItems"
End Sub

Private Sub Class_Initialize()
    Set Notification = New PropertyChangeNotification
    Set This.LineItems = New VBA.Collection
    Set This.NewLineItem = New OrderLineItemModel
    This.AddingNewItem = True
    This.ShipToBillingAddress = True
End Sub

Private Sub Class_Terminate()
    Set Notification = Nothing
    Set This.LineItems = Nothing
End Sub

Private Sub OnPropertyChanged(ByVal Name As String)
    INotifyPropertyChanged_OnPropertyChanged Me, Name
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal Name As String)
    Notification.Notify Source, Name
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    Notification.AddHandler Handler
End Sub

Public Function IsValid() As Boolean
    IsValid = This.LineItems.Count > 0 And Len(This.BillTo.Name) > 0 And Len(This.BillTo.Line1) > 0 And (This.ShipToBillingAddress Or (Len(This.ShipTo.Name) > 0 And Len(This.ShipTo.Line1) > 0))
End Function
