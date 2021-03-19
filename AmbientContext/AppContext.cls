VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder "AmbientContext"
'@PredeclaredId
Option Explicit
Implements IAppContext

Private Type TState
    'Abstract dependencies
    Factory As IAppContextFactory
    Current As IAppContext
    TimerProvider As ITimerProvider
    
    'Context-local state (previously global)
    Target As Range
    Properties As Scripting.Dictionary
    '...
    
End Type

Private This As TState

'@Description "Gets the current (or default) context."
Public Property Get Current() As IAppContext
Attribute Current.VB_Description = "Gets the current (or default) context."
    Errors.GuardNonDefaultInstance Me, AppContext, TypeName(Me)
    
    If This.Current Is Nothing Then
        Set This.Current = This.Factory.Create
        Errors.GuardNullReference This.Factory, TypeName(Me), "IAppContextFactory.Create returned Nothing."
    End If
    
    Set Current = This.Current
End Property

'@Description "Gets or sets/injects the factory that creates IAppContext instances."
Public Property Get Factory() As IAppContextFactory
Attribute Factory.VB_Description = "Gets or sets/injects the factory that creates IAppContext instances."
    Errors.GuardNonDefaultInstance Me, AppContext, TypeName(Me)
    Set Factory = This.Factory
End Property

Public Property Set Factory(ByVal Value As IAppContextFactory)
    Errors.GuardNonDefaultInstance Me, AppContext, TypeName(Me)
    Set This.Factory = Value
End Property

'@Description "Gets or sets/injects the service that provides the IAppTimer implementation."
Public Property Get TimerProvider() As ITimerProvider
Attribute TimerProvider.VB_Description = "Gets or sets/injects the service that provides the IAppTimer implementation."
    Errors.GuardNonDefaultInstance Me, AppContext, TypeName(Me)
    Set TimerProvider = This.TimerProvider
End Property

Public Property Set TimerProvider(ByVal Value As ITimerProvider)
    Errors.GuardNonDefaultInstance Me, AppContext, TypeName(Me)
    Set This.TimerProvider = Value
End Property

'@Description "Gets an instance of an object that can schedule a macro."
'@Ignore ShadowedDeclaration: VBA.DateTime.Timer
Public Property Get Timer() As IAppTimer
Attribute Timer.VB_Description = "Gets an instance of an object that can schedule a macro."
    Set Timer = This.TimerProvider.Timer
End Property

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is AppContext
End Property

Public Sub Clear()
    Set This.Current = Nothing
    Set This.Target = Nothing
    If Not This.Properties Is Nothing Then
        Set This.Properties = New Scripting.Dictionary
    End If
End Sub

Private Sub Class_Initialize()
    If IsDefaultInstance Then
        'must initialize context with sensible defaults:
        Set This.Factory = New AppContextFactory
        Set This.TimerProvider = New TimerProvider
    Else
        Set This.Properties = New Scripting.Dictionary
        'we want all instances to have the same provider instance:
        Set This.TimerProvider = AppContext.TimerProvider
    End If
End Sub

Private Property Get IAppContext_Caller() As Range
    Set IAppContext_Caller = Application.Caller
End Property

Private Sub IAppContext_Clear()
    Clear
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
End Property

Private Property Get IAppContext_Target() As Range
    Set IAppContext_Target = This.Target
End Property

Private Property Get IAppContext_Timer() As IAppTimer
    Set IAppContext_Timer = This.TimerProvider.Timer
End Property
