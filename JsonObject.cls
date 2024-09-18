VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsonObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("JSON.Core")
Option Explicit
Implements IJson

Private Const ModuleName As String = "JsonObject"

Private Type TData
    Items As Collection
End Type

Private This As TData

Private Sub Class_Initialize()
    Const FunctionName As String = "Class_Initialize"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)

    Set This.Items = New Collection
End Sub

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
    
    IJson_DataType = JsonDataTypeObject
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
    
    Dim Output As String
    Dim Pair As Pair
    For Each Pair In This.Items
        If (Output <> vbNullString) Then
            Output = Output & ","
        End If
        Output = Output & """" & Pair.Key & """:" & Pair.Value.ToString
    Next
    IJson_ToString = "{" & Output & "}"
End Function

Public Function ToString() As String
    Const FunctionName As String = "ToString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)

    ToString = IJson_ToString
End Function

Public Function Value() As String
    Const FunctionName As String = "Value"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)

    Value = "[Object]"
End Function

Public Function Count() As Long
    Const FunctionName As String = "Count"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Count = This.Items.Count
End Function

'@DefaultMember
Public Function Items(ByVal Index As String) As Pair
Attribute Items.VB_UserMemId = 0
    Const FunctionName As String = "Items"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)

    Dim Key As String
    Key = Service.ToBase64(Index)

    Set Items = This.Items(Key)
End Function

'@Enumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Const FunctionName As String = "NewEnum"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Set NewEnum = This.Items.[_NewEnum]
End Function

Private Sub Parse(ByVal Stream As StringStream)
    Const FunctionName As String = "Parse"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Stream.EatCharacter "{"
    Stream.EatSpaces
    Do While Stream.PeekCharacter <> "}"
        Dim Pair As Pair
        Set Pair = ParseMember(Stream)
        
        This.Items.Add Pair, Service.ToBase64(Pair.Key)
        Stream.EatSpaces
        If (Stream.PeekCharacter = ",") Then
            Stream.EatCharacter ","
            Stream.EatSpaces
        End If
    Loop
    Stream.EatCharacter "}"
End Sub

Private Function ParseMember(ByVal Stream As StringStream) As Pair
    Const FunctionName As String = "ParseMember"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Json As JsonString
    Set Json = Service.CreateJsonString(Stream)
    
    Dim Name As String
    Name = Json.Value
    Stream.EatSpaces
    Stream.EatCharacter ":"
    Stream.EatSpaces
    
    Dim Value As IJson
    Set Value = Service.ParseValue(Stream)
    
    Dim Pair As Pair
    Set Pair = Service.CreatePair(Name, Value)
    Set ParseMember = Pair
End Function
