VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Service"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("JSON.Factories")
Option Explicit

Private Const ModuleName As String = "Service"

Friend Function CreateStringStream(ByVal Data As String) As StringStream
    Const FunctionName As String = "CreateStringStream"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Stream As StringStream
    Set Stream = New StringStream
    
    Stream.Create Data
    Set CreateStringStream = Stream
End Function

Friend Function CreateJsonNull(ByVal Stream As StringStream) As JsonNull
    Const FunctionName As String = "CreateJsonNull"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Instance As JsonNull
    Set Instance = New JsonNull
    
    Instance.Create Stream
    Set CreateJsonNull = Instance
End Function

Friend Function CreateJsonBoolean(ByVal Stream As StringStream) As JsonBoolean
    Const FunctionName As String = "CreateJsonBoolean"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Instance As JsonBoolean
    Set Instance = New JsonBoolean
    
    Instance.Create Stream
    Set CreateJsonBoolean = Instance
End Function

Friend Function CreateJsonNumber(ByVal Stream As StringStream) As JsonNumber
    Const FunctionName As String = "CreateJsonNumber"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Instance As JsonNumber
    Set Instance = New JsonNumber
    
    Instance.Create Stream
    Set CreateJsonNumber = Instance
End Function

Friend Function CreateJsonString(ByVal Stream As StringStream) As JsonString
    Const FunctionName As String = "CreateJsonString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Instance As JsonString
    Set Instance = New JsonString
    
    Instance.Create Stream
    Set CreateJsonString = Instance
End Function

Friend Function CreateJsonArray(ByVal Stream As StringStream) As JsonArray
    Const FunctionName As String = "CreateJsonArray"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Instance As JsonArray
    Set Instance = New JsonArray
    
    Instance.Create Stream
    Set CreateJsonArray = Instance
End Function

Friend Function CreateJsonObject(ByVal Stream As StringStream) As JsonObject
    Const FunctionName As String = "CreateJsonObject"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Instance As JsonObject
    Set Instance = New JsonObject
    
    Instance.Create Stream
    Set CreateJsonObject = Instance
End Function

Friend Function CreatePair(ByVal First As String, ByVal Second As IJson) As Pair
    Const FunctionName As String = "CreateJsonObject"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Instance As Pair
    Set Instance = New Pair
    
    Instance.Create First, Second
    Set CreatePair = Instance
End Function

Friend Function ParseValue(ByVal Stream As StringStream) As IJson
    Const FunctionName As String = "ParseValue"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Debug.Assert Not Stream.EOF     '// Stream must not be empty
    Select Case GetDataType(Stream)
    Case jsondatatypeArray
        Set ParseValue = CreateJsonArray(Stream)
    Case JsonDataTypeBoolean
        Set ParseValue = CreateJsonBoolean(Stream)
    Case JsonDataTypeNull
        Set ParseValue = CreateJsonNull(Stream)
    Case JsonDataTypeNumber
        Set ParseValue = CreateJsonNumber(Stream)
    Case JsonDataTypeObject
        Set ParseValue = CreateJsonObject(Stream)
    Case JsonDataTypeString
        Set ParseValue = CreateJsonString(Stream)
    Case Else
        Err.Raise JsonExceptionUnexpectedCharacter, ModuleName & "." & FunctionName, "Unexpected token"
    End Select
End Function

Private Function GetDataType(ByVal Stream As StringStream) As Long
    Const FunctionName As String = "GetDataType"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    If (Stream.PeekCharacter = "[") Then
        GetDataType = jsondatatypeArray
    ElseIf (Stream.Match("^null")) Then
        GetDataType = JsonDataTypeNull
    ElseIf (Stream.Match("^(?:-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?)")) Then
        GetDataType = JsonDataTypeNumber
    ElseIf (Stream.PeekCharacter = "{") Then
        GetDataType = JsonDataTypeObject
    ElseIf (Stream.PeekCharacter = """") Then
        GetDataType = JsonDataTypeString
    ElseIf (Stream.Match("^true|false")) Then
        GetDataType = JsonDataTypeBoolean
    Else
        GetDataType = -1
    End If
End Function

Friend Function ToBase64(ByVal Value As String) As String
    Const FunctionName As String = "ToBase64"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Data() As Byte
    Data = StrConv(Value, vbFromUnicode)
    
    Dim Document As Object      '// MSXML2.DOMDocument
    Set Document = VBA.CreateObject("MSXML2.DOMDocument")
    
    Dim Node As Object      '// MSXML2.IXMLDOMElement
    Set Node = Document.CreateElement("b64")
    
    Node.DataType = "bin.base64"
    Node.NodetypedValue = Data
    ToBase64 = Node.text
End Function

Friend Function GetAs(ByVal Item As IJson, ByVal DataType As JsonDataTypeEnum) As IJson
    Const FunctionName As String = "GetAs"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Select Case DataType
    Case JsonDataTypeEnum.jsondatatypeArray
        Dim JsonArray As JsonArray
        Set JsonArray = Item
        Set GetAs = JsonArray
    Case JsonDataTypeEnum.JsonDataTypeBoolean
        Dim JsonBoolean As JsonBoolean
        Set JsonBoolean = Item
        Set GetAs = JsonBoolean
    Case JsonDataTypeEnum.JsonDataTypeNull
        Dim JsonNull As JsonNull
        Set JsonNull = Item
        Set GetAs = JsonNull
    Case JsonDataTypeEnum.JsonDataTypeNumber
        Dim JsonNumber As JsonNumber
        Set JsonNumber = Item
        Set GetAs = JsonNumber
    Case JsonDataTypeEnum.JsonDataTypeObject
        Dim JsonObject As JsonObject
        Set JsonObject = Item
        Set GetAs = JsonObject
    Case JsonDataTypeEnum.JsonDataTypeString
        Dim JsonString As JsonString
        Set JsonString = Item
        Set GetAs = JsonString
    Case Else
        Err.Raise 9     '// Subscript out of range
    End Select
End Function
