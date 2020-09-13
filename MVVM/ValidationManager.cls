VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Implements validation helper methods for the ViewModel to consume."
'@Exposed
'@Folder MVVM.Infrastructure.Bindings.Validation
'@ModuleDescription "Implements validation helper methods for the ViewModel to consume."
'@PredeclaredId
Option Explicit
Implements IHandleValidationError
Implements INotifyPropertyChanged

Public Event PropertyChanged(ByVal Source As Object, ByVal PropertyName As String)

Private Type TState
    ValidationErrors As Dictionary
    IsValid As Boolean
    ThrowOnValidationError As Boolean
    Handlers As Collection
End Type

Private this As TState

Public Function Create() As ValidationManager
    Set Create = New ValidationManager
End Function

Public Property Get IsValid() As Boolean
    IsValid = this.ValidationErrors.Count = 0
End Property

Public Property Get ThrowOnValidationError() As Boolean
    ThrowOnValidationError = this.ThrowOnValidationError
End Property

Public Property Let ThrowOnValidationError(ByVal RHS As Boolean)
    this.ThrowOnValidationError = RHS
End Property

Private Sub NotifyListeners(ByVal Source As Object)
    OnPropertyChanged Source, "Message"
    OnPropertyChanged Source, "IsValid"
End Sub

Private Sub OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    RaiseEvent PropertyChanged(Source, PropertyName)
End Sub

Private Sub Class_Initialize()
    Set this.ValidationErrors = New Dictionary
    Set this.Handlers = New Collection
    this.IsValid = True
End Sub

Private Sub IHandleValidationError_ClearValidationErrors(ByVal PropertyName As String)
    If this.ValidationErrors.Exists(PropertyName) Then
        this.ValidationErrors.Remove PropertyName
    End If
    NotifyListeners Me
    If this.ValidationErrors.Count > 0 Then Debug.Print TypeName(Me) & ": " & this.ValidationErrors.Count & " validation error(s)."
End Sub

Private Property Get IHandleValidationError_IsValid() As Boolean
    IHandleValidationError_IsValid = this.ValidationErrors.Count = 0
End Property

Private Property Get IHandleValidationError_Message() As String
    If this.ValidationErrors.Count = 0 Then Exit Property
    IHandleValidationError_Message = IHandleValidationError_ValidationErrors(0, 1)
End Property

Private Sub IHandleValidationError_OnValidationError(ByVal Source As IPropertyBinding, ByVal Message As String)
    If Not this.ValidationErrors.Exists(Source.SourcePropertyPath) Then
        this.ValidationErrors.Add Source.SourcePropertyPath, Message
    End If
    NotifyListeners Source.Source
End Sub

Private Property Get IHandleValidationError_ValidationErrors() As Variant
    Dim result() As Variant
    ReDim result(0 To this.ValidationErrors.Count, 0 To 1)
    Dim i As Long
    For i = LBound(this.ValidationErrors.Keys) To UBound(this.ValidationErrors.Keys)
        result(i, 0) = this.ValidationErrors.Keys(i)
        result(i, 1) = this.ValidationErrors.Item(this.ValidationErrors.Keys(i))
    Next
    IHandleValidationError_ValidationErrors = result
End Property

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    OnPropertyChanged Source, PropertyName
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    this.Handlers.Add Handler
End Sub
