VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsonNull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("JSON.Core")
Option Explicit
Implements IJson

Private Const ModuleName As String = "JsonNull"

Friend Sub Create(ByVal Stream As StringStream)
    Const FunctionName As String = "Create"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Parse Stream
End Sub

Private Function IJson_DataType() As JsonDataTypeEnum
    Const FunctionName As String = "IJson_DataType"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    IJson_DataType = JsonDataTypeNull
End Function

Public Function DataType() As JsonDataTypeEnum
    Const FunctionName As String = "DataType"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    DataType = IJson_DataType
End Function

Private Function IJson_ToString() As String
    Const FunctionName As String = "IJson_ToString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    IJson_ToString = """null"""
End Function

Public Function ToString() As String
    Const FunctionName As String = "ToString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    ToString = IJson_ToString
End Function

'@DefaultMember
Public Function Value() As Variant
Attribute Value.VB_UserMemId = 0
    Const FunctionName As String = "Value"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Value = Null
End Function

Private Sub Parse(ByVal Stream As StringStream)
    Const FunctionName As String = "Parse"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    If (Stream.Match("^null.*")) Then
        Stream.EatString "null"
    Else
        Err.Raise JsonExceptionUnexpectedToken, ModuleName & "." & FunctionName, "Token ""null"" expected."
    End If
End Sub
