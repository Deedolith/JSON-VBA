VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Base class for data writers."
'@ModuleDescription "Base class for data writers."
'@Exposed
'@Interface
'@Folder "JSON.Providers"
Option Explicit

'@Description "Data writing interface"
Public Sub Execute(ByRef Data As String)
Attribute Execute.VB_Description = "Data writing interface"
End Sub
