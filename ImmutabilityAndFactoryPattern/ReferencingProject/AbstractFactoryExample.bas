Attribute VB_Name = "AbstractFactoryExample"
'@Folder("VBAProject")
Option Explicit

Public Sub DoSomething()
    ManufactureSomeCar New AbstractCarFactory
End Sub

Private Sub ManufactureSomeCar(ByVal factory As ICarFactory)
    Dim myCar As ICar '<~ ManufactureSomeCar isn't coupled with any specific ICar implementation
    Set myCar = factory.Create(2016, "Civic", "Honda") '<~ ManufactureSomeCar doesn't need to know/care what specific implementation it's getting
    
    MsgBox "We have a " & myCar.Make & " " & myCar.Manufacturer & " " & myCar.Model & " here."
    'these assignments are illegal here, code won't compile if they're uncommented:
    'myCar.Make = 2014
    'myCar.Model = "Fit"
End Sub
