VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAppContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Encapsulates the data and macro invocation mechanism for a side-effecting UDF."
'@Folder "AmbientContext.Abstract"
'@ModuleDescription "Encapsulates the data and macro invocation mechanism for a side-effecting UDF."
'@Interface
Option Explicit

'@Description "Gets the cell that invoked the currently-running user-defined function (UDF), if applicable; Nothing otherwise."
Public Property Get Caller() As Range
Attribute Caller.VB_Description = "Gets the cell that invoked the currently-running user-defined function (UDF), if applicable; Nothing otherwise."
End Property

'@Description "Gets or sets the target reference cell that the side-effecting macro shall use."
Public Property Get Target() As Range
Attribute Target.VB_Description = "Gets or sets the target reference cell that the side-effecting macro shall use."
End Property

Public Property Set Target(ByVal Value As Range)
End Property

'@Description "Gets or sets a named value representing data passed between the UDF and the side-effecting macro."
Public Property Get Property(ByVal Name As String) As Variant
Attribute Property.VB_Description = "Gets or sets a named value representing data passed between the UDF and the side-effecting macro."
End Property

Public Property Let Property(ByVal Name As String, ByVal Value As Variant)
End Property

'@Description "Gets an array of all property names."
Public Property Get Properties() As Variant
Attribute Properties.VB_Description = "Gets an array of all property names."
End Property

'@Description "Gets an object that can schedule the execution of a macro."
Public Property Get Timer() As IAppTimer
Attribute Timer.VB_Description = "Gets an object that can schedule the execution of a macro."
End Property

'@Description "Clears all held state."
Public Sub Clear()
Attribute Clear.VB_Description = "Clears all held state."
End Sub
