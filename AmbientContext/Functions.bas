Attribute VB_Name = "Functions"
'@Folder "UDF"
Option Explicit

'@Description "A parameterized UDF demonstrating how Ambient Context can be leveraged to achieve testability with a side-effecting UDF."
Public Function TestUDF(ByVal SomeParameter As Double) As Boolean
Attribute TestUDF.VB_Description = "A parameterized UDF demonstrating how Ambient Context can be leveraged to achieve testability with a side-effecting UDF."
    On Error GoTo CleanFail
    
    With AppContext.Current
        
        Set .Target = .Caller.Offset(RowOffset:=1)
        .Property("Test1") = 42
        .Property("Test2") = 4.25 * SomeParameter
        .Timer.ExecuteMacroAsync
        
    End With
    
    TestUDF = True
CleanExit:
    Exit Function
CleanFail:
    TestUDF = False
    Resume CleanExit
    Resume
End Function

'@Description "Another UDF demonstrating how Ambient Context can be leveraged to achieve testability with a side-effecting UDF."
Public Function AnotherUDF() As Boolean
Attribute AnotherUDF.VB_Description = "Another UDF demonstrating how Ambient Context can be leveraged to achieve testability with a side-effecting UDF."
    On Error GoTo CleanFail
    
    With AppContext.Current
        Set .Target = .Caller.Offset(RowOffset:=1)
        .Property("Test1") = DateTime.Now
        .Timer.ExecuteMacroAsync
    End With
    
    AnotherUDF = True
CleanExit:
    Exit Function
CleanFail:
    AnotherUDF = False
    Resume CleanExit
    Resume
End Function
