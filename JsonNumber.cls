VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsonNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("JSON.Core")
Option Explicit
Implements IJson

Private Const ModuleName As String = "JsonNumber"

Private Type TData
    Value As Double
End Type

Private This As TData

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
    
    IJson_DataType = JsonDataTypeNumber
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
    
    IJson_ToString = Replace(CStr(This.Value), ",", ".")
End Function

Public Function ToString() As String
    Const FunctionName As String = "ToString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    ToString = IJson_ToString
End Function

'@DefaultMember
Public Function Value() As Double
Attribute Value.VB_UserMemId = 0
    Const FunctionName As String = "Value"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Value = This.Value
End Function

Private Sub Parse(ByVal Stream As StringStream)
    Const FunctionName As String = "Parse"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Const Pattern As String = "^(?:-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?)"
    
    If (Stream.Match(Pattern)) Then
        Dim Token As String
        Token = Stream.PeekString(Pattern)

        On Error GoTo RegionalSettingsError
            This.Value = CDbl(Token)
        On Error GoTo 0
        Stream.EatString Token
    Else
        Err.Raise JsonExceptionUnexpectedToken, ModuleName & "." & FunctionName, "Numeric token expected."
    End If
Exit Sub
RegionalSettingsError:
    Select Case Err.Number
    Case 13     '// Type mismatch due to regional settings
        Err.Clear
        This.Value = CDbl(Replace$(Token, ".", ","))
        Resume Next
    Case Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End Select
End Sub
