VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AggregateErrorFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A composite error formatter implementation."
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder MVVM.Infrastructure.Bindings.Validation.UI
'@ModuleDescription "A composite error formatter implementation."
'@PredeclaredId
'@Exposed
Option Explicit
Implements IValidationErrorFormatter

Private Type TState
    Context As Object
    Formatters As Collection
End Type

Private this As TState

Public Function Create(ByVal Context As IViewModel, ParamArray Formatters() As Variant) As IValidationErrorFormatter
    GuardClauses.GuardNonDefaultInstance Me, AggregateErrorFormatter, TypeName(Me)
    GuardClauses.GuardExpression UBound(Formatters) < 0, TypeName(Me), "At least one formatter is required."

    Dim result As AggregateErrorFormatter
    Set result = New AggregateErrorFormatter
    Set result.Context = Context
    
    Dim i As Long
    For i = LBound(Formatters) To UBound(Formatters)
        
        If Not IsEmpty(Formatters(i)) And Not Formatters(i) Is Nothing Then
            Dim Formatter As IValidationErrorFormatter
            On Error Resume Next
            Set Formatter = Formatters(i)
            If Not Formatter Is Nothing Then result.AddFormatter Formatter
            On Error GoTo 0
        End If
        
    Next
    
    Set Create = result
    
End Function

Public Sub AddFormatter(ByVal Formatter As IValidationErrorFormatter)
    GuardClauses.GuardExpression IsDefaultInstance, TypeName(Me)
    GuardClauses.GuardNullReference Formatter
    
    this.Formatters.Add Formatter

End Sub

Public Property Get Context() As Object
    Set Context = this.Context
End Property

Public Property Set Context(ByVal RHS As Object)
    GuardClauses.GuardDoubleInitialization this.Context, TypeName(Me)
    GuardClauses.GuardExpression IsDefaultInstance, TypeName(Me)
    GuardClauses.GuardNullReference RHS
    
    Set this.Context = RHS

End Property

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is AggregateErrorFormatter
End Property

Private Sub Class_Initialize()
    If Not IsDefaultInstance Then Set this.Formatters = New Collection
End Sub

Private Sub IValidationErrorFormatter_Apply(ByVal PropertyName As String, ByVal Message As String)
    Dim Formatter As IValidationErrorFormatter
    For Each Formatter In this.Formatters
        Formatter.Apply PropertyName, Message
    Next
End Sub

Private Sub IValidationErrorFormatter_Restore()
    Dim Formatter As IValidationErrorFormatter
    For Each Formatter In this.Formatters
        Formatter.Restore
    Next
End Sub

Private Property Get IValidationErrorFormatter_Target() As Object
    Set IValidationErrorFormatter_Target = this.Context
End Property
