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

Private Const ModuleName As String = "Services"

'@Description "Instanciate a new StringStream object"
Friend Function CreateStringStream(ByVal Data As String) As JSON.StringStream
Attribute CreateStringStream.VB_Description = "Instanciate a new StringStream object"
#If DEV Then
    Const FunctionName As String = "CreateStringStream"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim SS As JSON.StringStream
    Set SS = New JSON.StringStream

    SS.Create Data
    Set CreateStringStream = SS
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

    Dim rx As Object        '// VBScript_RegExp_55.RegExp
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.MultiLine = True

    rx.Pattern = "\\([""\\\/])"                     '// unescape double quotes, anti-slash and slash
    Value = rx.Replace(Value, "$1")
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

    Dim rx As Object        '// VBScript_RegExp_55.RegExp
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.MultiLine = True
    rx.Pattern = "\\u([a-fA-F0-9]{4})"
    
    Dim Matchs As Object        '// VBScript_RegExp_55.MatchCollection
    Set Matchs = rx.Execute(Data)
    
    Dim Value As String
    Value = Data
    Dim Match As Object     '// VBScript_RegExp_55.Match
    For Each Match In Matchs
        rx.Pattern = "(\" & Match.Value & ")"
        Value = rx.Replace(Value, ChrW$(Val("&H" & Match.SubMatches(0))))
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
        
    Dim rx As Object        '// VBScript_RegExp_55.RegExp
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.MultiLine = True

    Dim Value As String
    rx.Pattern = "([""\\\/])"       '// escape double quotes, anti-slash and slash
    Value = rx.Replace(Data, "\$1")
    rx.Pattern = vbLf               '// escape new line
    Value = rx.Replace(Value, "\n")
    rx.Pattern = vbCr               '// escape carriage return
    Value = rx.Replace(Value, "\r")
    rx.Pattern = vbBack             '// escape backspace
    Value = rx.Replace(Value, "\b")
    rx.Pattern = vbTab              '// escape tab
    Value = rx.Replace(Value, "\t")
    rx.Pattern = vbVerticalTab      '// escape vertical tab
    Value = rx.Replace(Value, "\f")
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
    
    Dim XML As Object        '// MSXML2.DOMDocument
    Set XML = VBA.CreateObject("MSXML2.DOMDocument")
    
    Dim Node As Object       '// MSXML2.IXMLDOMElement
    Set Node = XML.CreateElement("b64")
    
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
Friend Function CreateBoolean(ByRef SS As JSON.StringStream) As JSON.JBoolean
Attribute CreateBoolean.VB_Description = "Instanciate a new JBoolean object from a stream"
    Debug.Assert Not (SS Is Nothing)

#If DEV Then
    Const FunctionName As String = "CreateBoolean"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim JBool As JSON.JBoolean
    Set JBool = New JSON.JBoolean

    JBool.Create SS
    Set CreateBoolean = JBool
End Function

'@Description "Instanciate a new JNull object from a stream"
Friend Function CreateNull(ByRef SS As JSON.StringStream) As JSON.JNull
Attribute CreateNull.VB_Description = "Instanciate a new JNull object from a stream"
        Debug.Assert Not (SS Is Nothing)

#If DEV Then
    Const FunctionName As String = "CreateNull"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim JNull As JSON.JNull
    Set JNull = New JSON.JNull
    
    JNull.Create SS
    Set CreateNull = JNull
End Function

'@Description "Instaciate a new JNumber object from a stream"
Friend Function CreateNumber(ByRef SS As JSON.StringStream) As JSON.JNumber
Attribute CreateNumber.VB_Description = "Instaciate a new JNumber object from a stream"
    Debug.Assert Not (SS Is Nothing)

#If DEV Then
    Const FunctionName As String = "CreateNumber"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Number As JSON.JNumber
    Set Number = New JSON.JNumber

    Number.Create SS
    Set CreateNumber = Number
End Function

'@Description "Instanciate a new JString object from a stream"
Friend Function CreateString(ByRef SS As JSON.StringStream) As JSON.JString
Attribute CreateString.VB_Description = "Instanciate a new JString object from a stream"
    Debug.Assert Not (SS Is Nothing)

#If DEV Then
    Const FunctionName As String = "CreateString"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim JString As JSON.JString
    Set JString = New JSON.JString

    JString.Create SS
    Set CreateString = JString
End Function

'@Description "Instanciate a new JArray object from a stream"
Friend Function CreateArray(ByRef SS As JSON.StringStream) As JSON.Jarray
Attribute CreateArray.VB_Description = "Instanciate a new JArray object from a stream"
    Debug.Assert Not (SS Is Nothing)    '// SS must be a valid instance
    Debug.Assert Not (SS.EOF)           '// SS must not be empty

#If DEV Then
    Const FunctionName As String = "CreateArray"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If
    
    Dim Table As JSON.Jarray
    Set Table = New JSON.Jarray

    Table.Create SS
    Set CreateArray = Table
End Function
