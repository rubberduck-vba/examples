Attribute VB_Name = "Example"
'@Folder MVVM.Example
Option Explicit

'@Description "Runs the MVVM example UI."
Public Sub Run()
Attribute Run.VB_Description = "Runs the MVVM example UI."
'here a more elaborate application would wire-up dependencies for complex commands,
'and then property-inject them into the ViewModel via a factory method e.g. SomeViewModel.Create(args).

    Dim ViewModel As ExampleViewModel
    Set ViewModel = ExampleViewModel.Create
    
    'TODO: implement and inject an ExampleCommand; CommandButton1 will automatically bind it.
    
    'ViewModel properties can be set before or after it's wired to the View.
    'ViewModel.SourcePath = "TEST"
    ViewModel.SomeOption = True
    
    Set ViewModel.SomeCommand = New BrowseCommand
    
    Dim View As IView
    Set View = ExampleView.Create(ViewModel, BindingManager.Create)
    
    If View.ShowDialog Then
        Debug.Print ViewModel.SourcePath, ViewModel.SomeOption, ViewModel.SomeOtherOption
    Else
        Debug.Print "Dialog was cancelled."
    End If
    
End Sub
