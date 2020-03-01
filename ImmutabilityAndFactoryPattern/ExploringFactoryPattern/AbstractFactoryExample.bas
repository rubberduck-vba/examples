Attribute VB_Name = "AbstractFactoryExample"
'@Folder("Examples")
Option Explicit

Public Sub DoSomething()
    
    ' let's make a factory that creates brand new cars whose Manufacturer is "Honda"
    Dim factory As ISimplerCarFactory
    Set factory = ManufacturerCarFactory.Create("Honda")
    
    ' let's make a bunch of cars
    Dim cars As Collection
    Set cars = CreateSomeHondaCars(factory)
    
    ' ...and now consume them
    ListAllCars cars
    
End Sub

Private Function CreateSomeHondaCars(ByVal factory As ISimplerCarFactory) As Collection
'NOTE: this function doesn't know or care what specific ISimplerCarFactory it's working with.
    Dim cars As Collection
    Set cars = New Collection
    cars.Add factory.Create("Civic")
    cars.Add factory.Create("Accord")
    cars.Add factory.Create("CRV")
    Set CreateSomeCars = Collection
End Function

Private Sub ListAllCars(ByVal cars As Collection)
'NOTE: this procedure doesn't know or care whas specific ICar implementation it's working with.
    Dim c As ICar
    For Each c In cars
        Debug.Print c.Make, c.Manufacturer, c.Model
    Next
End Sub
