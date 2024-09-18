VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder("JSON.DataReaders")
Option Explicit
Implements IReader

Private Const ModuleName As String = "StringReader"

Private Type TData
    Data As String
End Type

Private This As TData

Friend Sub Create(ByVal Data As String)
    Const FunctionName As String = "Create"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    This.Data = Data
End Sub

Private Function IReader_Execute() As String
    Const FunctionName As String = "IReader_Execute"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    IReader_Execute = This.Data
End Function
