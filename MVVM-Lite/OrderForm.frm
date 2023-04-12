VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OrderForm 
   Caption         =   "Rubberduck Swag Shop"
   ClientHeight    =   8100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10680
   OleObjectBlob   =   "OrderForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OrderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TView
    Result As Boolean
    OrderModel As OrderHeaderModel
    ViewModel As INotifyPropertyChanged
    Bindings As VBA.Collection
    AddLineItemCommand As ICommand
    ClearLineItemsCommand As ICommand
End Type

Private This As TView

Public Property Get Model() As OrderHeaderModel
    Set Model = This.OrderModel
End Property

Public Property Get IsConfirmed() As Boolean
    IsConfirmed = This.Result
End Property

Private Sub AddLineItemButton_Click()
    This.AddLineItemCommand.Execute This.OrderModel
End Sub

Private Sub CancelButton_Click()
    Me.Hide
End Sub

Private Sub ClearOrderItemsButton_Click()
    This.ClearLineItemsCommand.Execute This.OrderModel
End Sub

Private Sub ConfirmButton_Click()
    This.Result = True
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    
    Set This.Bindings = New VBA.Collection
    Set This.OrderModel = New OrderHeaderModel
    Set This.ViewModel = This.OrderModel
    
    Set This.AddLineItemCommand = New AddLineItemCommand
    Set This.ClearLineItemsCommand = New ClearLineItemsCommand
    
    ConfigureBindings This.OrderModel
    
End Sub

Private Sub ConfigureBindings(ByVal Model As INotifyPropertyChanged)

    Const EnabledProperty As String = "Enabled"
    Const ListProperty As String = "List"
    
    This.Bindings.Add BindTextBox(Me.BillToNameBox, "BillToName", This.OrderModel)
    This.Bindings.Add BindTextBox(Me.BillToAddressLine1, "BillToLine1", This.OrderModel)
    This.Bindings.Add BindTextBox(Me.BillToAddressLine2, "BillToLine2", This.OrderModel)
    This.Bindings.Add BindTextBox(Me.BillToAddressLine3, "BillToLine3", This.OrderModel)
    This.Bindings.Add BindTextBox(Me.BillToEmailBox, "EmailAddress", This.OrderModel)
    This.Bindings.Add BindCheckBox(Me.BillToContributorBox, "IsContributor", This.OrderModel)
    
    This.Bindings.Add BindCheckBox(Me.ShipToSameBox, "ShipToBillingAddress", This.OrderModel)
    This.Bindings.Add BindTextBox(Me.ShipToNameBox, "ShipToName", This.OrderModel)
    This.Bindings.Add BindTextBox(Me.ShipToAddressLine1, "ShipToLine1", This.OrderModel)
    This.Bindings.Add BindTextBox(Me.ShipToAddressLine2, "ShipToLine2", This.OrderModel)
    This.Bindings.Add BindTextBox(Me.ShipToAddressLine3, "ShipToLine3", This.OrderModel)
    
    This.Bindings.Add BindProperty(Me.ShipToAddressLabel, EnabledProperty, "ShipToBillingAddress", This.OrderModel, InvertBoolean:=True)
    This.Bindings.Add BindProperty(Me.ShipToNameLabel, EnabledProperty, "ShipToBillingAddress", This.OrderModel, InvertBoolean:=True)
    This.Bindings.Add BindProperty(Me.ShipToNameBox, EnabledProperty, "ShipToBillingAddress", This.OrderModel, InvertBoolean:=True)
    This.Bindings.Add BindProperty(Me.ShipToAddressLine1, EnabledProperty, "ShipToBillingAddress", This.OrderModel, InvertBoolean:=True)
    This.Bindings.Add BindProperty(Me.ShipToAddressLine2, EnabledProperty, "ShipToBillingAddress", This.OrderModel, InvertBoolean:=True)
    This.Bindings.Add BindProperty(Me.ShipToAddressLine3, EnabledProperty, "ShipToBillingAddress", This.OrderModel, InvertBoolean:=True)
    
    This.Bindings.Add BindProperty(Me.ItemSkuSelectBox, ListProperty, "Value", InventorySheet.Table.ListColumns("SKU").DataBodyRange)
    This.Bindings.Add BindComboBox(Me.ItemSkuSelectBox, "SKU", This.OrderModel.NewLineItem)
    This.Bindings.Add BindTextBox(Me.ItemQuantityBox, "Quantity", This.OrderModel.NewLineItem)
    This.Bindings.Add BindTextBox(Me.ItemPriceBox, "Price", This.OrderModel.NewLineItem)
    
    This.Bindings.Add BindProperty(Me.LineItemsList, ListProperty, "LineItems", This.OrderModel)

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub
