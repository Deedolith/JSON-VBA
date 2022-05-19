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
'@Folder "JSON.Services"
Option Explicit

#If DEV Then
    Private Const ModuleName As String = "Services"
#End If

'@Description "Instanciate a new StringStream object"
Friend Function CreateStringStream(ByVal Data As String) As JSON.StringStream
Attribute CreateStringStream.VB_Description = "Instanciate a new StringStream object"
#If DEV Then
    Const FunctionName As String = "CreateStringStream"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim StringStream As JSON.StringStream
    Set StringStream = New JSON.StringStream

    StringStream.Create Data
    Set CreateStringStream = StringStream
End Function

'@Description "Unescape strings"
Public Function Unescape(ByVal Data As String) As String
Attribute Unescape.VB_Description = "Unescape strings"
#If DEV Then
    Const FunctionName As String = "Unescape"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

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
#If DEV Then
    Const FunctionName As String = "UnescapeUnicode"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

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
#If DEV Then
    Const FunctionName As String = "Escape"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If
        
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
#If DEV Then
    Const FunctionName As String = "EscapeUnicode"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

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
#If DEV Then
    Const FunctionName As String = "ToBase64"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

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

#If DEV Then
    Public Property Get LibraryName() As String
        LibraryName = "JSON"
    End Property
#End If

#If DEV Then
    Public Function CreateLogger(ByVal ModuleName As String, ByVal FunctionName As String) As Logger
        Dim Logger As JSON.Logger
        Set Logger = New Logger
        Logger.Create ModuleName, FunctionName
        Set CreateLogger = Logger
    End Function
#End If

'@Description "Instanciate a new JBoolean object from a stream"
Friend Function CreateBoolean(ByVal StringStream As JSON.StringStream) As JSON.JBoolean
Attribute CreateBoolean.VB_Description = "Instanciate a new JBoolean object from a stream"
    Debug.Assert Not (StringStream Is Nothing)

#If DEV Then
    Const FunctionName As String = "CreateBoolean"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim JBool As JSON.JBoolean
    Set JBool = New JSON.JBoolean

    JBool.Create StringStream
    Set CreateBoolean = JBool
End Function

'@Description "Instanciate a new JNull object from a stream"
Friend Function CreateNull(ByVal StringStream As JSON.StringStream) As JSON.JNull
Attribute CreateNull.VB_Description = "Instanciate a new JNull object from a stream"
        Debug.Assert Not (StringStream Is Nothing)

#If DEV Then
    Const FunctionName As String = "CreateNull"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim JNull As JSON.JNull
    Set JNull = New JSON.JNull
    
    JNull.Create StringStream
    Set CreateNull = JNull
End Function

'@Description "Instaciate a new JNumber object from a stream"
Friend Function CreateNumber(ByVal StringStream As JSON.StringStream) As JSON.JNumber
Attribute CreateNumber.VB_Description = "Instaciate a new JNumber object from a stream"
    Debug.Assert Not (StringStream Is Nothing)

#If DEV Then
    Const FunctionName As String = "CreateNumber"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Number As JSON.JNumber
    Set Number = New JSON.JNumber

    Number.Create StringStream
    Set CreateNumber = Number
End Function

'@Description "Instanciate a new JString object from a stream"
Friend Function CreateString(ByVal StringStream As JSON.StringStream) As JSON.JString
Attribute CreateString.VB_Description = "Instanciate a new JString object from a stream"
    Debug.Assert Not (StringStream Is Nothing)

#If DEV Then
    Const FunctionName As String = "CreateString"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim JString As JSON.JString
    Set JString = New JSON.JString

    JString.Create StringStream
    Set CreateString = JString
End Function

'@Description "Instanciate a new JArray object from a stream"
Friend Function CreateArray(ByVal StringStream As JSON.StringStream) As JSON.Jarray
Attribute CreateArray.VB_Description = "Instanciate a new JArray object from a stream"
    Debug.Assert Not (StringStream Is Nothing)    '// SS must be a valid instance
    Debug.Assert Not (StringStream.EOF)           '// SS must not be empty

#If DEV Then
    Const FunctionName As String = "CreateArray"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If
    
    Dim Jarray As JSON.Jarray
    Set Jarray = New JSON.Jarray

    Jarray.Create StringStream
    Set CreateArray = Jarray
End Function

'@Description "Convert a generic JSON object to the requested type"
Friend Function GetValueAs(ByVal Value As Object, ByVal DataType As JSON.JType) As Object
Attribute GetValueAs.VB_Description = "Convert a generic JSON object to the requested type"
#If DEV Then
    Const FunctionName As String = "GetValueAs"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Select Case DataType
    Case JSON.JType.JSArray
        Dim Jarray As JSON.Jarray
        Set Jarray = Value
        Set GetValueAs = Jarray
    Case JSON.JSBoolean
        Dim JBoolean As JSON.JBoolean
        Set JBoolean = Value
        Set GetValueAs = JBoolean
    Case JSON.JType.JSNull
        Dim JNull As JSON.JNull
        Set JNull = Value
        Set GetValueAs = JNull
    Case JSON.JType.JSNumber
        Dim JNumber As JSON.JNumber
        Set JNumber = Value
        Set GetValueAs = JNumber
    Case JSON.JType.JSObject
        Dim JObject As JSON.JObject
        Set JObject = Value
        Set GetValueAs = JObject
    Case JSON.JType.JSString
        Dim JString As JSON.JString
        Set JString = Value
        Set GetValueAs = JString
    Case Else
        Err.Raise 9     '// subscript out of range
    End Select
End Function

'@Description "Parse a JSON value and return the corresponding object"
Friend Function ParseValue(ByVal StringStream As JSON.StringStream) As Object
Attribute ParseValue.VB_Description = "Parse a JSON value and return the corresponding object"
    Debug.Assert Not (StringStream Is Nothing)    '// SS must be a valid instance
    Debug.Assert Not (StringStream.EOF)           '// SS must not be empty

#If DEV Then
    Const FunctionName As String = "ParseValue"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

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
        Err.Raise JSON.JException.JUnexpectedCharacter, "Services.ParseValue", "Unexpected token"
    End If
End Function

'@Description "Check if the next value as an array"
Private Function IsArray(ByVal StringStream As JSON.StringStream) As Boolean
Attribute IsArray.VB_Description = "Check if the next value as an array"
    Debug.Assert Not (StringStream Is Nothing)

#If DEV Then
    Const FunctionName As String = "IsArray"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    IsArray = StringStream.PeekCharacter = "["
End Function

'@Description "Check if the next value is a boolean"
Private Function IsBoolean(ByVal StringStream As JSON.StringStream) As Boolean
Attribute IsBoolean.VB_Description = "Check if the next value is a boolean"
    Debug.Assert Not (StringStream Is Nothing)

#If DEV Then
    Const FunctionName As String = "IsBoolean"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    IsBoolean = StringStream.GetStringFromRegEx("^(?:true|false)") <> vbNullString
End Function

'@Description "Check if the next value is a null"
Private Function IsNull(ByVal StringStream As JSON.StringStream) As Boolean
Attribute IsNull.VB_Description = "Check if the next value is a null"
    Debug.Assert Not (StringStream Is Nothing)

#If DEV Then
    Const FunctionName As String = "IsNull"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    IsNull = StringStream.GetStringFromRegEx("^null") <> vbNullString
End Function

'@Description "Check if the next value is a number"
Private Function IsNumber(ByVal StringStream As JSON.StringStream) As Boolean
Attribute IsNumber.VB_Description = "Check if the next value is a number"
    Debug.Assert Not (StringStream Is Nothing)

#If DEV Then
    Const FunctionName As String = "IsNumber"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    IsNumber = StringStream.GetStringFromRegEx("^(?:-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?)") <> vbNullString
End Function

'@Description "Check if the next value is an object"
Private Function IsObject(ByVal StringStream As JSON.StringStream) As Boolean
Attribute IsObject.VB_Description = "Check if the next value is an object"
    Debug.Assert Not (StringStream Is Nothing)

#If DEV Then
    Const FunctionName As String = "IsObject"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    IsObject = StringStream.PeekCharacter = "{"
End Function

'@Description "Check if the next value is a string"
Private Function IsString(ByVal StringStream As JSON.StringStream) As Boolean
Attribute IsString.VB_Description = "Check if the next value is a string"
    Debug.Assert Not (StringStream Is Nothing)

#If DEV Then
    Const FunctionName As String = "IsString"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    IsString = StringStream.PeekCharacter = """"
End Function

'@Description "Instanciate a new Pair object"
Friend Function CreatePair(ByVal Name As String, ByVal Element As Object) As JSON.Pair
Attribute CreatePair.VB_Description = "Instanciate a new Pair object"
#If DEV Then
    Const FunctionName As String = "CreatePair"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    If ((TypeOf Element Is Jarray) Or (TypeOf Element Is JBoolean) Or (TypeOf Element Is JNull) Or (TypeOf Element Is JNumber) Or (TypeOf Element Is JString) Or (TypeOf Element Is JObject)) Then
        Dim Pair As JSON.Pair
        Set Pair = New JSON.Pair
        Pair.Create Name, Element
        Set CreatePair = Pair
    Else
        Err.Raise 13        '// Type Mismatch
    End If
End Function

'@Description "Instanciate a new JObject object from a stream"
Friend Function CreateObject(ByVal StringStream As JSON.StringStream) As JSON.JObject
Attribute CreateObject.VB_Description = "Instanciate a new JObject object from a stream"
    Debug.Assert Not (StringStream Is Nothing)    '// SS must be a valid instance
    Debug.Assert Not (StringStream.EOF)           '// SS must not be empty

#If DEV Then
    Const FunctionName As String = "CreateObject"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If
    
    Dim JObject As JSON.JObject
    Set JObject = New JSON.JObject

    JObject.Create StringStream
    Set CreateObject = JObject
End Function
