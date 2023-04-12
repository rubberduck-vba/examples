VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CreateInvoiceCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICommand

Private Function ICommand_CanExecute(ByVal Parameter As Object) As Boolean
    If TypeOf Parameter Is OrderHeaderModel Then
        Dim Model As OrderHeaderModel
        Set Model = Parameter
        ICommand_CanExecute = Model.IsValid
    End If
End Function

Private Sub ICommand_Execute(ByVal Parameter As Object)
    If Not TypeOf Parameter Is OrderHeaderModel Then Err.Raise 5
    
    Dim Model As OrderHeaderModel
    Set Model = Parameter
    
    Dim InvoiceNumber As String
    InvoiceNumber = ThisWorkbook.Names("NextInvoiceNumber").RefersToRange.Value
    
    Dim InvoiceDate As Date
    InvoiceDate = VBA.DateTime.Date
    
    InvoicesSheet.AddInvoice InvoiceNumber, InvoiceDate, Model
    
    Dim LineItem As OrderLineItemModel
    Dim ItemIndex As Long
    
    InvoiceTemplateSheet.SetInvoiceHeading InvoiceNumber, InvoiceDate
    
    For Each LineItem In Model.LineItemObjects
        ItemIndex = ItemIndex + 1
        SalesSheet.AddLineItem InvoiceNumber, ItemIndex, InvoiceDate, LineItem
        InvoiceTemplateSheet.SetLineItem ItemIndex, LineItem.SKU, LineItem.Quantity, LineItem.Price * LineItem.Quantity
    Next
    
    InvoiceTemplateSheet.SetBillTo Model.BillToName, Model.BillToLine1, Model.BillToLine2, Model.BillToLine3
    InvoiceTemplateSheet.SetShipTo Model.ShipToName, Model.ShipToLine1, Model.ShipToLine2, Model.ShipToLine3
    InvoiceTemplateSheet.SetShipping "US Air", 15
    InvoiceTemplateSheet.PrintPreview
    
End Sub
