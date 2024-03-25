VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Base class for data readers."
'@ModuleDescription "Base class for data readers."
'@Exposed
'@Interface
'@Folder "JSON.Providers"
Option Explicit

'@Description "Data reading interface"
Public Function Execute() As String
Attribute Execute.VB_Description = "Data reading interface"
End Function
