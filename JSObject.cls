VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Interface
'@Folder("JSON")
Option Explicit

Public Function ToJSONString() As String: End Function
Public Function ToString(Optional ByVal IndentMultiplier As Long = 0) As String: End Function
