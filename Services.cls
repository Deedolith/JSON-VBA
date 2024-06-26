VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Services"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@IgnoreModule SuspiciousPredeclaredInstanceAccess, AssignmentNotUsed, VariableNotUsed
'@Folder "Services"
Option Explicit

Private Const ModuleName As String = "Services"

'@Description "Instanciate a new StringStream JSObject"
Friend Function CreateStringStream(ByVal Data As String) As StringStream
Attribute CreateStringStream.VB_Description = "Instanciate a new StringStream JSObject"
    Const FunctionName As String = "CreateStringStream"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim StringStream As StringStream
    Set StringStream = New StringStream

    StringStream.Create Data
    Set CreateStringStream = StringStream
End Function

'@Description "Unescape strings"
Public Function Unescape(ByVal Data As String) As String
Attribute Unescape.VB_Description = "Unescape strings"
    Const FunctionName As String = "Unescape"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim Value As String
    Value = UnescapeUnicode(Data)                    '// unescape unicodes

    Dim RegEx As Object        '// VBScript_RegExp_55.RegExp
    Set RegEx = VBA.CreateObject("VBScript.RegExp")
    RegEx.Global = True
    RegEx.MultiLine = True

    RegEx.Pattern = "\\([""\\\/])"                     '// unescape double quotes, anti-slash and slash
    Value = RegEx.Replace(Value, "$1")
    Value = Replace(Value, "\b", vbBack)            '// unescape backspace
    Value = Replace(Value, "\f", vbVerticalTab)     '// unescape vertical tab
    Value = Replace(Value, "\n", vbLf)              '// unescape new line
    Value = Replace(Value, "\r", vbCr)              '// unescape carriage return
    Value = Replace(Value, "\t", vbTab)             '// unescape tab
    Unescape = Value
End Function

'@Description "Unescape unicode characters"
Private Function UnescapeUnicode(ByVal Data As String) As String
Attribute UnescapeUnicode.VB_Description = "Unescape unicode characters"
    Const FunctionName As String = "UnescapeUnicode"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim RegEx As Object        '// VBScript_RegExp_55.RegExp
    Set RegEx = VBA.CreateObject("VBScript.RegExp")
    RegEx.Global = True
    RegEx.MultiLine = True
    RegEx.Pattern = "\\u([a-fA-F0-9]{4})"
    
    Dim Matchs As Object        '// VBScript_RegExp_55.MatchCollection
    Set Matchs = RegEx.Execute(Data)
    
    Dim Value As String
    Value = Data
    Dim Match As Object     '// VBScript_RegExp_55.Match
    For Each Match In Matchs
        RegEx.Pattern = "(\" & Match.Value & ")"
        Value = RegEx.Replace(Value, ChrW$(Val("&H" & Match.SubMatches(0))))
    Next
    UnescapeUnicode = Value
End Function

'@Description "Escape string"
Public Function Escape(ByVal Data As String) As String
Attribute Escape.VB_Description = "Escape string"
    Const FunctionName As String = "Escape"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)
        
    Dim RegEx As Object        '// VBScript_RegExp_55.RegExp
    Set RegEx = VBA.CreateObject("VBScript.RegExp")
    RegEx.Global = True
    RegEx.MultiLine = True

    Dim Value As String
    RegEx.Pattern = "([""\\\/])"       '// escape double quotes, anti-slash and slash
    Value = RegEx.Replace(Data, "\$1")
    RegEx.Pattern = vbLf               '// escape new line
    Value = RegEx.Replace(Value, "\n")
    RegEx.Pattern = vbCr               '// escape carriage return
    Value = RegEx.Replace(Value, "\r")
    RegEx.Pattern = vbBack             '// escape backspace
    Value = RegEx.Replace(Value, "\b")
    RegEx.Pattern = vbTab              '// escape tab
    Value = RegEx.Replace(Value, "\t")
    RegEx.Pattern = vbVerticalTab      '// escape vertical tab
    Value = RegEx.Replace(Value, "\f")
    Value = EscapeUnicode(Value)
    Escape = Value
End Function

'@Description "Escape unicode characters"
Private Function EscapeUnicode(ByVal Data As String) As String
Attribute EscapeUnicode.VB_Description = "Escape unicode characters"
    Const FunctionName As String = "EscapeUnicode"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim Value As String
    Value = vbNullString
    
    Dim i As Long
    For i = 1 To Len(Data)
        Dim Character As String
        Character = Mid$(Data, i, 1)
        Select Case AscW(Character)
                '// don't touch ASCII characters
        Case 32 To 126
            Value = Value & Character
        Case Else
            Value = Value & "\u" & Right$("0000" & Hex$(AscW(Character)), 4)
        End Select
    Next
    EscapeUnicode = Value
End Function

'@Description "Encode a string to base64"
Friend Function ToBase64(ByVal text As String) As String
Attribute ToBase64.VB_Description = "Encode a string to base64"
    Const FunctionName As String = "ToBase64"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim Data() As Byte
    Data = StrConv(text, vbFromUnicode)
    
    Dim XMLDocument As Object        '// MSXML2.DOMDocument
    Set XMLDocument = VBA.CreateObject("MSXML2.DOMDocument")
    
    Dim Node As Object       '// MSXML2.IXMLDOMElement
    Set Node = XMLDocument.CreateElement("b64")
    
    Node.DataType = "bin.base64"
    Node.nodeTypedValue = Data
    ToBase64 = Node.text
End Function

Public Function CreateErrorLogger(ByVal ModuleName As String, ByVal FunctionName As String) As ErrorLogger
#If DebugMode Then
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = New ErrorLogger

    ErrorLogger.Create ModuleName, FunctionName
    Set CreateErrorLogger = ErrorLogger
#End If
End Function

'@Description "Instanciate a new JBoolean JSObject from a stream"
Friend Function CreateBoolean(ByVal StringStream As StringStream) As JBoolean
Attribute CreateBoolean.VB_Description = "Instanciate a new JBoolean JSObject from a stream"
    Debug.Assert Not (StringStream Is Nothing)

    Const FunctionName As String = "CreateBoolean"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim JBool As JBoolean
    Set JBool = New JBoolean

    JBool.Create StringStream
    Set CreateBoolean = JBool
End Function

'@Description "Instanciate a new JNull JSObject from a stream"
Friend Function CreateNull(ByVal StringStream As StringStream) As JNull
Attribute CreateNull.VB_Description = "Instanciate a new JNull JSObject from a stream"
        Debug.Assert Not (StringStream Is Nothing)

    Const FunctionName As String = "CreateNull"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim JNull As JNull
    Set JNull = New JNull
    
    JNull.Create StringStream
    Set CreateNull = JNull
End Function

'@Description "Instaciate a new JNumber JSObject from a stream"
Friend Function CreateNumber(ByVal StringStream As StringStream) As JNumber
Attribute CreateNumber.VB_Description = "Instaciate a new JNumber JSObject from a stream"
    Debug.Assert Not (StringStream Is Nothing)

    Const FunctionName As String = "CreateNumber"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim Number As JNumber
    Set Number = New JNumber

    Number.Create StringStream
    Set CreateNumber = Number
End Function

'@Description "Instanciate a new JString JSObject from a stream"
Friend Function CreateString(ByVal StringStream As StringStream) As JString
Attribute CreateString.VB_Description = "Instanciate a new JString JSObject from a stream"
    Debug.Assert Not (StringStream Is Nothing)

    Const FunctionName As String = "CreateString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim JString As JString
    Set JString = New JString

    JString.Create StringStream
    Set CreateString = JString
End Function

'@Description "Instanciate a new JArray JSObject from a stream"
Friend Function CreateArray(ByVal StringStream As StringStream) As JArray
Attribute CreateArray.VB_Description = "Instanciate a new JArray JSObject from a stream"
    Debug.Assert Not (StringStream Is Nothing)    '// SS must be a valid instance
    Debug.Assert Not (StringStream.EOF)           '// SS must not be empty

    Const FunctionName As String = "CreateArray"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim JArray As JArray
    Set JArray = New JArray

    JArray.Create StringStream
    Set CreateArray = JArray
End Function

'@Description "Convert a generic JSON JSObject to the requested type"
Friend Function GetValueAs(ByVal Value As JSObject, ByVal DataType As JType) As JSObject
Attribute GetValueAs.VB_Description = "Convert a generic JSON JSObject to the requested type"
    Const FunctionName As String = "GetValueAs"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Select Case DataType
    Case JType.JSArray
        Dim JArray As JArray
        Set JArray = Value
        Set GetValueAs = JArray
    Case JSBoolean
        Dim JBoolean As JBoolean
        Set JBoolean = Value
        Set GetValueAs = JBoolean
    Case JType.JSNull
        Dim JNull As JNull
        Set JNull = Value
        Set GetValueAs = JNull
    Case JType.JSNumber
        Dim JNumber As JNumber
        Set JNumber = Value
        Set GetValueAs = JNumber
    Case JType.JSObject
        Dim JObject As JObject
        Set JObject = Value
        Set GetValueAs = JObject
    Case JType.JSString
        Dim JString As JString
        Set JString = Value
        Set GetValueAs = JString
    Case Else
        Err.Raise 9     '// subscript out of range
    End Select
End Function

'@Description "Parse a JSON value and return the corresponding JSObject"
Friend Function ParseValue(ByVal StringStream As StringStream) As JSObject
Attribute ParseValue.VB_Description = "Parse a JSON value and return the corresponding JSObject"
    Debug.Assert Not (StringStream Is Nothing)    '// SS must be a valid instance
    Debug.Assert Not (StringStream.EOF)           '// SS must not be empty

    Const FunctionName As String = "ParseValue"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    StringStream.DiscardSpaces
    If (IsObject(StringStream)) Then
        Set ParseValue = CreateObject(StringStream)
    ElseIf (IsArray(StringStream)) Then
        Set ParseValue = CreateArray(StringStream)
    ElseIf (IsString(StringStream)) Then
        Set ParseValue = CreateString(StringStream)
    ElseIf (IsNumber(StringStream)) Then
        Set ParseValue = CreateNumber(StringStream)
    ElseIf (IsBoolean(StringStream)) Then
        Set ParseValue = CreateBoolean(StringStream)
    ElseIf (IsNull(StringStream)) Then
        Set ParseValue = CreateNull(StringStream)
    Else
        Err.Raise JException.JUnexpectedCharacter, "Services.ParseValue", "Unexpected token"
    End If
End Function

'@Description "Check if the next value as an array"
Private Function IsArray(ByVal StringStream As StringStream) As Boolean
Attribute IsArray.VB_Description = "Check if the next value as an array"
    Debug.Assert Not (StringStream Is Nothing)

    Const FunctionName As String = "IsArray"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    IsArray = StringStream.PeekCharacter = "["
End Function

'@Description "Check if the next value is a boolean"
Private Function IsBoolean(ByVal StringStream As StringStream) As Boolean
Attribute IsBoolean.VB_Description = "Check if the next value is a boolean"
    Debug.Assert Not (StringStream Is Nothing)

    Const FunctionName As String = "IsBoolean"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    IsBoolean = StringStream.GetStringFromRegEx("^(?:true|false)") <> vbNullString
End Function

'@Description "Check if the next value is a null"
Private Function IsNull(ByVal StringStream As StringStream) As Boolean
Attribute IsNull.VB_Description = "Check if the next value is a null"
    Debug.Assert Not (StringStream Is Nothing)

    Const FunctionName As String = "IsNull"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    IsNull = StringStream.GetStringFromRegEx("^null") <> vbNullString
End Function

'@Description "Check if the next value is a number"
Private Function IsNumber(ByVal StringStream As StringStream) As Boolean
Attribute IsNumber.VB_Description = "Check if the next value is a number"
    Debug.Assert Not (StringStream Is Nothing)

    Const FunctionName As String = "IsNumber"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    IsNumber = StringStream.GetStringFromRegEx("^(?:-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?)") <> vbNullString
End Function

'@Description "Check if the next value is an JSObject"
Private Function IsObject(ByVal StringStream As StringStream) As Boolean
Attribute IsObject.VB_Description = "Check if the next value is an JSObject"
    Debug.Assert Not (StringStream Is Nothing)

    Const FunctionName As String = "IsObject"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    IsObject = StringStream.PeekCharacter = "{"
End Function

'@Description "Check if the next value is a string"
Private Function IsString(ByVal StringStream As StringStream) As Boolean
Attribute IsString.VB_Description = "Check if the next value is a string"
    Debug.Assert Not (StringStream Is Nothing)

    Const FunctionName As String = "IsString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    IsString = StringStream.PeekCharacter = DOUBLEQUOTE
End Function

'@Description "Instanciate a new Pair JSObject"
Friend Function CreatePair(ByVal Name As String, ByVal Element As JSObject) As Pair
Attribute CreatePair.VB_Description = "Instanciate a new Pair JSObject"
    Const FunctionName As String = "CreatePair"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    If ((TypeOf Element Is JArray) Or (TypeOf Element Is JBoolean) Or (TypeOf Element Is JNull) Or (TypeOf Element Is JNumber) Or (TypeOf Element Is JString) Or (TypeOf Element Is JObject)) Then
        Dim Pair As Pair
        Set Pair = New Pair
        Pair.Create Name, Element
        Set CreatePair = Pair
    Else
        Err.Raise 13        '// Type Mismatch
    End If
End Function

'@Description "Instanciate a new JObject JSObject from a stream"
Friend Function CreateObject(ByVal StringStream As StringStream) As JObject
Attribute CreateObject.VB_Description = "Instanciate a new JObject JSObject from a stream"
    Debug.Assert Not (StringStream Is Nothing)    '// SS must be a valid instance
    Debug.Assert Not (StringStream.EOF)           '// SS must not be empty

    Const FunctionName As String = "CreateObject"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim JObject As JObject
    Set JObject = New JObject

    JObject.Create StringStream
    Set CreateObject = JObject
End Function
