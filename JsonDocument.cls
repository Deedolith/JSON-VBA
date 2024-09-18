VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsonDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder("JSON.Core")
Option Explicit

Private Const ModuleName As String = "Document"

Private Type TData
    Value As IJson
End Type

Private This As TData

Public Sub LoadFrom(ByVal Reader As IReader)
On Error GoTo Error
    Const FunctionName As String = "LoadFrom"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Data As String
    Data = Reader.Execute
    
    Dim Length As Long
    Length = Len(Data)
    
    Dim Stream As StringStream
    Set Stream = Service.CreateStringStream(Data)
    
    Stream.EatSpaces
    Set This.Value = Service.ParseValue(Stream)
Exit Sub
Error:
    Err.Raise Err.Number, Err.Source, Err.Description & " at position " & Length - Len(Stream.Value) & " (" & Left$(Stream.Value, 20) & ")."
End Sub

Public Sub WriteTo(ByVal Writer As IWriter)
    Const FunctionName As String = "WriteTo"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Writer.Execute This.Value.ToString
End Sub

Friend Function GetAs(ByVal DataType As JsonDataTypeEnum) As IJson
    Const FunctionName As String = "GetAs"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Set GetAs = Service.GetAs(This.Value, DataType)
End Function

Public Function Query(ByVal XPath As String) As IJson
    Const FunctionName As String = "Query"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Path As Collection
    Set Path = ParseXPath(XPath)
    
    Dim Value As IJson
    Set Value = This.Value
    
    Dim Item As Variant
    For Each Item In Path
        If (IsNumeric(Item)) Then
            Dim JsonArray As JsonArray
            Set JsonArray = Service.GetAs(Value, jsondatatypeArray)
            Set Value = JsonArray.Items(CLng(Item))
        Else
            Dim JsonObject As JsonObject
            Set JsonObject = Service.GetAs(Value, JsonDataTypeObject)
            Set Value = JsonObject.Items(Item).Value
        End If
    Next
    Set Query = Value
End Function

Private Function ParseXPath(ByVal XPath As String) As Collection
    Const FunctionName As String = "ParseXPath"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Output As Collection
    Set Output = New Collection
    
    Dim Stream As StringStream
    Set Stream = Service.CreateStringStream(XPath)
    
    Stream.EatCharacter "/"
    
    Dim Character As String
    Do While Not Stream.EOF
        Character = Stream.PeekCharacter
        If (Character = "[") Then
            Stream.EatCharacter "["
            
            Dim Number As String
            Number = ParseNumber(Stream)
            Output.Add CInt(Number)
            Stream.EatCharacter "]"
        ElseIf (Character <> "/") Then
            Dim Data As String
            Data = ParseString(Stream)
            Output.Add Data
        Else
            Stream.EatCharacter "/"
        End If
    Loop
    Set ParseXPath = Output
End Function

Private Function ParseNumber(ByVal Stream As StringStream) As String
    Const FunctionName As String = "ParseNumber"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Number As String
    Number = Stream.PeekString("^[1-9]\d*")
    Stream.EatString Number
    ParseNumber = Number
End Function

Private Function ParseString(ByVal Stream As StringStream) As String
    Const FunctionName As String = "ParseString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Data As String
    Data = Stream.PeekString("^[^\[/]+")
    Stream.EatString Data
    ParseString = Data
End Function
