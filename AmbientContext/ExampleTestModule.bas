Attribute VB_Name = "ExampleTestModule"
'@IgnoreModule ShadowedDeclaration
'@Folder "Tests"
'@TestModule
Option Explicit
Option Private Module

#Const LateBind = LateBindTests 'precompiler constant defined in Project Properties

#If LateBind Then
    Private Assert As Object
'    Private Fakes As Object
#Else
    Private Assert As Rubberduck.AssertClass
'    Private Fakes As Rubberduck.FakesProvider
#End If

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.AssertClass")
'        Set Fakes = CreateObject("Rubberduck.FakesProvider")
    #Else
        Set Assert = New Rubberduck.AssertClass
'        Set Fakes = New Rubberduck.FakesProvider
    #End If
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
'    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    AppContext.Clear
End Sub

'@TestCleanup
Private Sub TestCleanup()
    AppContext.Clear
End Sub

'@TestMethod("Infrastructure")
Private Sub TestUDF_InvokesContextCaller()
    'inject the test factory:
    Set AppContext.Factory = New TestContextFactory
    
    'get the test context:
    Dim Context As TestContext
    Set Context = AppContext.Current
    
    'run the UDF:
    Dim Result As Boolean
    Result = Functions.TestUDF(0)
    
    'Assert that the UDF has invoked the IAppContext.Caller getter once:
    Const Expected As Long = 1
    Assert.AreEqual Expected, Context.CallerInvokes, "IAppContext.Caller was invoked " & Context.CallerInvokes & " times; expected " & Expected
End Sub

'@TestMethod("AppData")
Private Sub TestUDF_SetsTarget()
    'inject the test factory:
    Set AppContext.Factory = New TestContextFactory
    
    'get the test context:
    Dim Context As TestContext
    Set Context = AppContext.Current
    
    'run the UDF:
    Dim Result As Boolean
    Result = Functions.TestUDF(0)
    
    'Assert that the UDF has invoked the IAppContext.Target setter once:
    Const Expected As Long = 1
    Assert.AreEqual Expected, Context.TargetSetterInvokes, "IAppContext.Target setter was invoked " & Context.TargetSetterInvokes & " times; expected " & Expected
End Sub

'@TestMethod("Infrastructure")
Private Sub TestUDF_SchedulesMacro()
    'inject the test factory:
    Set AppContext.Factory = New TestContextFactory
    
    'get the test context:
    Dim Context As TestContext
    Set Context = AppContext.Current
    
    'stub the timer:
    Dim StubTimer As TestTimer
    Set StubTimer = AppContext.Current.Timer
    
    'run the UDF:
    Dim Result As Boolean
    Result = Functions.TestUDF(0)
    
    'Assert that the UDF has invoked IAppContext.ScheduleMacro once:
    Const Expected As Long = 1
    Assert.AreEqual Expected, StubTimer.ExecuteMacroAsyncInvokes, "IAppTimer.ExecuteMacroAsync was invoked " & StubTimer.ExecuteMacroAsyncInvokes & " times; expected " & Expected
End Sub

'@TestMethod("Infrastructure")
Private Sub AnotherUDF_InvokesContextCaller()
    'inject the test factory:
    Set AppContext.Factory = New TestContextFactory
    
    'get the test context:
    Dim Context As TestContext
    Set Context = AppContext.Current
    
    'run the UDF:
    Dim Result As Boolean
    Result = Functions.AnotherUDF
    
    'Assert that the UDF has invoked the IAppContext.Caller getter once:
    Const Expected As Long = 1
    Assert.AreEqual Expected, Context.CallerInvokes, "IAppContext.Caller was invoked " & Context.CallerInvokes & " times; expected " & Expected
End Sub

'@TestMethod("AppData")
Private Sub AnotherUDF_SetsTarget()
    'inject the test factory:
    Set AppContext.Factory = New TestContextFactory
    
    'get the test context:
    Dim Context As TestContext
    Set Context = AppContext.Current
    
    'run the UDF:
    Dim Result As Boolean
    Result = Functions.AnotherUDF
    
    'Assert that the UDF has invoked the IAppContext.Target setter once:
    Const Expected As Long = 1
    Assert.AreEqual Expected, Context.TargetSetterInvokes, "IAppContext.Target setter was invoked " & Context.TargetSetterInvokes & " times; expected " & Expected
End Sub

'@TestMethod("Infrastructure")
Private Sub AnotherUDF_SchedulesMacro()
    'inject the test factory:
    Set AppContext.Factory = New TestContextFactory
    
    'get the test context:
    Dim Context As TestContext
    Set Context = AppContext.Current
    
    'stub the timer:
    Dim StubTimer As TestTimer
    Set StubTimer = AppContext.Current.Timer
    
    'run the UDF:
    Dim Result As Boolean
    Result = Functions.AnotherUDF
    
    'Assert that the UDF has invoked IAppContext.ScheduleMacro once:
    Const Expected As Long = 1
    Assert.AreEqual Expected, StubTimer.ExecuteMacroAsyncInvokes, "IAppTimer.ExecuteMacroAsync was invoked " & StubTimer.ExecuteMacroAsyncInvokes & " times; expected " & Expected
End Sub

