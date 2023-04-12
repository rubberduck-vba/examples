VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AddLineItemCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICommand

Private Function ICommand_CanExecute(ByVal Parameter As Object) As Boolean
    ICommand_CanExecute = TypeOf Parameter Is OrderHeaderModel
End Function

Private Sub ICommand_Execute(ByVal Parameter As Object)
    If Not TypeOf Parameter Is OrderHeaderModel Then Err.Raise 5
    
    Dim Model As OrderHeaderModel
    Set Model = Parameter
    
    Dim Item As OrderLineItemModel
    Set Item = New OrderLineItemModel
    
    Item.SKU = Model.NewLineItem.SKU
    Item.Quantity = Model.NewLineItem.Quantity
    Item.Price = Model.NewLineItem.Price
    
    Model.AddLineItem Item
    
    Model.NewLineItem.SKU = vbNullString
    Model.NewLineItem.Quantity = 1
    Model.NewLineItem.Price = 0
    
End Sub
