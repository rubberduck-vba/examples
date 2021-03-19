VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "An implementation of IAppContext that can be injected in place of AppContext for testing."
'@IgnoreModule ProcedureNotUsed
'@Folder "Tests.Stubs"
'@ModuleDescription "An implementation of IAppContext that can be injected in place of AppContext for testing."
Option Explicit
Implements IAppContext

Private Type TContext
    MacroName As String
    Target As Range
    Properties As Scripting.Dictionary
    Factory As IAppContextFactory
    Timer As IAppTimer
    
    FactorySetterInvokes As Long
    TargetSetterInvokes As Long
    
    CallerInvokes As Long

End Type

Private This As TContext

Public Property Get CallerInvokes() As Long
    CallerInvokes = This.CallerInvokes
End Property

Public Property Get FactorySetterInvokes() As Long
    FactorySetterInvokes = This.FactorySetterInvokes
End Property

Public Property Get TargetSetterInvokes() As Long
    TargetSetterInvokes = This.TargetSetterInvokes
End Property

Public Property Get Target() As Range
    Set Target = This.Target
End Property

Public Property Set Target(ByVal Value As Range)
    Set This.Target = Value
End Property

Public Property Get Property(ByVal Name As String) As Variant
    Property = This.Properties.Item(Name)
End Property

Public Property Let Property(ByVal Name As String, ByVal Value As Variant)
    This.Properties.Item(Name) = Value
End Property

Public Property Set Property(ByVal Name As String, ByVal Value As Object)
    Set This.Properties.Item(Name) = Value
End Property

Private Sub Class_Initialize()
    Set This.Factory = New TestContextFactory
    Set This.Timer = New TestTimer
    Set This.Properties = New Scripting.Dictionary
End Sub

Private Property Get IAppContext_Caller() As Range
    This.CallerInvokes = This.CallerInvokes + 1
    Set IAppContext_Caller = HostSheet.Cells(1, 1)
End Property

Private Sub IAppContext_Clear()
    Set This.Target = Nothing
    If Not This.Properties Is Nothing Then
        Set This.Properties = New Scripting.Dictionary
    End If
End Sub

Private Property Get IAppContext_Properties() As Variant
    IAppContext_Properties = This.Properties.Keys
End Property

Private Property Let IAppContext_Property(ByVal Name As String, ByVal Value As Variant)
    This.Properties.Item(Name) = Value
End Property

Private Property Set IAppContext_Property(ByVal Name As String, ByVal Value As Object)
    Set This.Properties.Item(Name) = Value
End Property

Private Property Get IAppContext_Property(ByVal Name As String) As Variant
    IAppContext_Property = This.Properties.Item(Name)
End Property

Private Property Set IAppContext_Target(ByVal Value As Range)
    Set This.Target = Value
    This.TargetSetterInvokes = This.TargetSetterInvokes + 1
End Property

Private Property Get IAppContext_Target() As Range
    Set IAppContext_Target = This.Target
End Property

Private Property Set IAppContext_Timer(ByVal Value As IAppTimer)
    Set This.Timer = Value
End Property

Private Property Get IAppContext_Timer() As IAppTimer
    Set IAppContext_Timer = This.Timer
End Property
