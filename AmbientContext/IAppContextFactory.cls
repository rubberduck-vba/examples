VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAppContextFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "An abstract factory that creates IAppContext instances."
'@Folder "AmbientContext.Abstract"
'@ModuleDescription "An abstract factory that creates IAppContext instances."
'@Interface
Option Explicit

'@Description "Creates and returns an AppContext instance."
Public Function Create() As IAppContext
Attribute Create.VB_Description = "Creates and returns an AppContext instance."
End Function
