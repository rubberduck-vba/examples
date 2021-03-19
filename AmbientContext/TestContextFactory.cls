VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestContextFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A factory that creates TestContext instances."
'@Folder "Tests.Stubs"
'@ModuleDescription "A factory that creates TestContext instances."
Option Explicit
Implements IAppContextFactory

Private Function IAppContextFactory_Create() As IAppContext
    Set IAppContextFactory_Create = New TestContext
End Function
