VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationErrorFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object that describes how a control changes its appearance given a validation error."
'@Folder MVVM.Infrastructure.Bindings.Validation.UI
'@ModuleDescription "An object that describes how a control changes its appearance given a validation error."
'@PredeclaredId
'@Exposed
Option Explicit
Implements IValidationErrorFormatter

Private Const DefaultErrorBackColor As Long = &HC0C0FF
Private Const DefaultErrorBorderColor As Long = &HC0&
Private Const DefaultErrorForeColor As Long = &HC0&

Private Type TState
    Target As Object
    Applied As Boolean
    
    FormatBackgroundColor As Boolean
    FormatBorderColor As Boolean
    FormatForeColor As Boolean
    FormatFontBold As Boolean
    
    InitialBackgroundColor As Long
    InitialBorderColor As Long
    InitialForeColor As Long
    InitialFontBold As Boolean
    
    ErrorBackgroundColor As Long
    ErrorBorderColor As Long
    ErrorForeColor As Long
    ErrorFontBold As Boolean
    
    TargetOnlyVisibleOnError As Boolean
    TriggerPropertyName As String
    
End Type

Private this As TState

Public Function Create(ByVal Target As Object) As ValidationErrorFormatter
    Dim result As ValidationErrorFormatter
    Set result = New ValidationErrorFormatter
    Set result.Target = Target
    Set Create = result
End Function

Public Function WithErrorBackgroundColor(Optional ByVal Color As Variant = DefaultErrorBackColor) As ValidationErrorFormatter
    this.ErrorBackgroundColor = Color
    this.FormatBackgroundColor = True
    Set WithErrorBackgroundColor = Me
End Function

Public Function WithErrorBorderColor(Optional ByVal Color As Variant = DefaultErrorBorderColor) As ValidationErrorFormatter
    this.ErrorBorderColor = Color
    this.FormatBorderColor = True
    Set WithErrorBorderColor = Me
End Function

Public Function WithErrorForeColor(Optional ByVal Color As Variant = DefaultErrorForeColor) As ValidationErrorFormatter
    this.ErrorForeColor = Color
    this.FormatForeColor = True
    Set WithErrorForeColor = Me
End Function

Public Function WithErrorBoldFont() As ValidationErrorFormatter
    this.FormatFontBold = True
    this.ErrorFontBold = True
    Set WithErrorBoldFont = Me
End Function

Public Function WithTargetOnlyVisibleOnError(Optional ByVal PropertyName As String) As ValidationErrorFormatter
    this.TargetOnlyVisibleOnError = True
    this.TriggerPropertyName = PropertyName
    Set WithTargetOnlyVisibleOnError = Me
End Function

Public Property Get Target() As Object
    Set Target = this.Target
End Property

Public Property Set Target(ByVal RHS As Object)
    
    Select Case True
    
        Case TypeOf RHS Is MSForms.TextBox
            Dim TextBoxTarget As MSForms.TextBox
            Set TextBoxTarget = RHS
            this.InitialBackgroundColor = TextBoxTarget.BackColor
            this.InitialBorderColor = TextBoxTarget.BorderColor
            this.InitialFontBold = TextBoxTarget.Font.Bold
            this.InitialForeColor = TextBoxTarget.ForeColor
            
            Set this.Target = TextBoxTarget
            
        Case TypeOf RHS Is MSForms.CheckBox 'NOTE: MSForms.OptionButton also matches this interface
            Dim CheckBoxTarget As MSForms.CheckBox
            Set CheckBoxTarget = RHS
            this.InitialFontBold = CheckBoxTarget.Font.Bold
            this.InitialForeColor = CheckBoxTarget.ForeColor
            
            Set this.Target = CheckBoxTarget
            
        Case TypeOf RHS Is MSForms.Label
            Dim LabelTarget As MSForms.Label
            Set LabelTarget = RHS
            this.InitialBackgroundColor = LabelTarget.BackColor
            this.InitialBorderColor = LabelTarget.BorderColor
            this.InitialFontBold = LabelTarget.Font.Bold
            this.InitialForeColor = LabelTarget.ForeColor
            
            Set this.Target = LabelTarget
            
        Case TypeOf RHS Is MSForms.Frame
            Dim FrameTarget As MSForms.Frame
            Set FrameTarget = RHS
            this.InitialBackgroundColor = FrameTarget.BackColor
            this.InitialBorderColor = FrameTarget.BorderColor
            this.InitialFontBold = FrameTarget.Font.Bold
            this.InitialForeColor = FrameTarget.ForeColor
        
            Set this.Target = FrameTarget
            
        Case TypeOf RHS Is MSForms.Image
            Dim ImageTarget As MSForms.Image
            Set ImageTarget = RHS
            this.InitialBackgroundColor = ImageTarget.BackColor
            this.InitialBorderColor = ImageTarget.BorderColor
            
            Set this.Target = ImageTarget
        
        Case Else
            Debug.Print TypeName(Me) & ": Target of type " & TypeName(RHS) & " is not currently supported."
            
    End Select
    
End Property

Private Sub IValidationErrorFormatter_Apply(ByVal PropertyName As String, ByVal Message As String)
    If this.Applied Then Exit Sub
    
    Dim Control As MSForms.Control
    Set Control = this.Target
    If Not this.TargetOnlyVisibleOnError Then
        Control.ControlTipText = Message
    End If
    
    Select Case True
    
        Case TypeOf this.Target Is MSForms.TextBox
            Dim TextBoxTarget As MSForms.TextBox
            Set TextBoxTarget = this.Target
            If this.FormatBackgroundColor Then TextBoxTarget.BackColor = this.ErrorBackgroundColor
            If this.FormatBorderColor Then TextBoxTarget.BorderColor = this.ErrorBorderColor
            If this.FormatFontBold Then TextBoxTarget.Font.Bold = this.ErrorFontBold
            If this.FormatForeColor Then TextBoxTarget.ForeColor = this.ErrorForeColor
            this.Applied = True
            
        Case TypeOf this.Target Is MSForms.CheckBox 'NOTE: MSForms.OptionButton also matches this interface
            Dim CheckBoxTarget As MSForms.CheckBox
            Set CheckBoxTarget = this.Target
            If this.FormatFontBold Then CheckBoxTarget.Font.Bold = this.ErrorFontBold
            If this.FormatForeColor Then CheckBoxTarget.ForeColor = this.ErrorForeColor
            this.Applied = True
            
        Case TypeOf this.Target Is MSForms.Label
            Dim LabelTarget As MSForms.Label
            Set LabelTarget = this.Target
            If this.FormatBackgroundColor Then LabelTarget.BackColor = this.ErrorBackgroundColor
            If this.FormatBorderColor Then LabelTarget.BorderColor = this.ErrorBorderColor
            If this.FormatFontBold Then LabelTarget.Font.Bold = this.ErrorFontBold
            If this.FormatForeColor Then LabelTarget.ForeColor = this.ErrorForeColor
            If this.TargetOnlyVisibleOnError Then LabelTarget.Caption = Message
            this.Applied = True
            
        Case TypeOf this.Target Is MSForms.Frame
            Dim FrameTarget As MSForms.Frame
            Set FrameTarget = this.Target
            If this.FormatBackgroundColor Then FrameTarget.BackColor = this.ErrorBackgroundColor
            If this.FormatBorderColor Then FrameTarget.BorderColor = this.ErrorBorderColor
            If this.FormatFontBold Then FrameTarget.Font.Bold = this.ErrorFontBold
            If this.FormatForeColor Then FrameTarget.ForeColor = this.ErrorForeColor
            this.Applied = True
            
        Case TypeOf this.Target Is MSForms.Image
            Dim ImageTarget As MSForms.Image
            Set ImageTarget = this.Target
            If this.FormatBackgroundColor Then ImageTarget.BackColor = this.ErrorBackgroundColor
            If this.FormatBorderColor Then ImageTarget.BorderColor = this.ErrorBorderColor
            this.Applied = True
            
    End Select
    
    If this.TargetOnlyVisibleOnError _
    And (this.TriggerPropertyName = vbNullString Or this.TriggerPropertyName = PropertyName) Then
        Control.Visible = True
    End If

End Sub

Private Sub IValidationErrorFormatter_Restore()
    If Not this.Applied Then Exit Sub

    Dim Control As MSForms.Control
    Set Control = this.Target
    Control.ControlTipText = vbNullString
    If this.TargetOnlyVisibleOnError Then Control.Visible = False
    
    Select Case True
    
        Case TypeOf this.Target Is MSForms.TextBox
            Dim TextBoxTarget As MSForms.TextBox
            Set TextBoxTarget = this.Target
            If this.FormatBackgroundColor Then TextBoxTarget.BackColor = this.InitialBackgroundColor
            If this.FormatBorderColor Then TextBoxTarget.BorderColor = this.InitialBorderColor
            If this.FormatFontBold Then TextBoxTarget.Font.Bold = this.InitialFontBold
            If this.FormatForeColor Then TextBoxTarget.ForeColor = this.InitialForeColor
            
            this.Applied = False
            
        Case TypeOf this.Target Is MSForms.CheckBox 'NOTE: MSForms.OptionButton also matches this interface
            Dim CheckBoxTarget As MSForms.CheckBox
            Set CheckBoxTarget = this.Target
            If this.FormatFontBold Then CheckBoxTarget.Font.Bold = this.InitialFontBold
            If this.FormatForeColor Then CheckBoxTarget.ForeColor = this.InitialForeColor
            
            this.Applied = False
            
        Case TypeOf this.Target Is MSForms.Label
            Dim LabelTarget As MSForms.Label
            Set LabelTarget = this.Target
            If this.FormatBackgroundColor Then LabelTarget.BackColor = this.InitialBackgroundColor
            If this.FormatBorderColor Then LabelTarget.BorderColor = this.InitialBorderColor
            If this.FormatFontBold Then LabelTarget.Font.Bold = this.InitialFontBold
            If this.FormatForeColor Then LabelTarget.ForeColor = this.InitialForeColor
            
            this.Applied = False
            
        Case TypeOf this.Target Is MSForms.Frame
            Dim FrameTarget As MSForms.Frame
            Set FrameTarget = this.Target
            If this.FormatBackgroundColor Then FrameTarget.BackColor = this.InitialBackgroundColor
            If this.FormatBorderColor Then FrameTarget.BorderColor = this.InitialBorderColor
            If this.FormatFontBold Then FrameTarget.Font.Bold = this.InitialFontBold
            If this.FormatForeColor Then FrameTarget.ForeColor = this.InitialForeColor
        
            this.Applied = False
            
        Case TypeOf this.Target Is MSForms.Image
            Dim ImageTarget As MSForms.Image
            Set ImageTarget = this.Target
            If this.FormatBackgroundColor Then ImageTarget.BackColor = this.InitialBackgroundColor
            If this.FormatBorderColor Then ImageTarget.BorderColor = this.InitialBorderColor
            
            this.Applied = False
            
    End Select

End Sub

Private Property Get IValidationErrorFormatter_Target() As Object
    Set IValidationErrorFormatter_Target = this.Target
End Property

