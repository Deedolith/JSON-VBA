VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("JSON.Core")
Option Explicit

Private Const ModuleName As String = "Pair"

Private Type TData
    First As String
    Second As IJson
End Type

Private This As TData

Friend Sub Create(ByVal First As String, ByVal Second As IJson)
    Const FunctionName As String = "Create"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)

    This.First = First
    Set This.Second = Second
End Sub

Public Function Key() As String
    Const FunctionName As String = "Key"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)

    Key = This.First
End Function

'@DefaultMember
Public Function Value() As IJson
Attribute Value.VB_UserMemId = 0
    Const FunctionName As String = "Value"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)

    Set Value = This.Second
End Function
