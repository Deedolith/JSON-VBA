VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IJson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@IgnoreModule
'@Interface
'@Folder("JSON.Core")
Option Explicit

Public Function DataType() As JsonDataTypeEnum: End Function
Public Function ToString() As String: End Function
