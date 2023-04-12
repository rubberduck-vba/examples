Attribute VB_Name = "Invoicing"
Option Explicit

Public Sub NewCustomerOrder()

    Dim Model As OrderHeaderModel
    
    With New OrderForm
        .Show
        If .IsConfirmed Then
            Dim Command As ICommand
            Set Command = New CreateInvoiceCommand
            Command.Execute .Model
        End If
    End With
End Sub
