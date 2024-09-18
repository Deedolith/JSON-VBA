VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsonString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("JSON.Core")
Option Explicit
Implements IJson

Private Const ModuleName As String = "JsonString"

Private Type TData
    Value As String
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

    IJson_DataType = JsonDataTypeString
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

    IJson_ToString = """" & Escape(This.Value) & """"
End Function

Public Function ToString() As String
    Const FunctionName As String = "ToString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)

    ToString = IJson_ToString
End Function

'@DefaultMember
Public Function Value() As String
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
    
    Const Pattern As String = "^""(?:\\(?:[""\\\/bfnrt]|u[a-fA-F0-9]{4})|[^""\\\0-\x1F\x7F]+)*"""
    
    If (Stream.Match(Pattern)) Then
        Dim Token As String
        Token = Stream.PeekString(Pattern)
        This.Value = Unescape(Mid$(Token, 2, Len(Token) - 2))
        Stream.EatString Token
    Else
        Err.Raise JsonExceptionUnexpectedToken, ModuleName & "." & FunctionName, "String expected."
    End If
End Sub

Private Function Escape(ByVal Value As String) As String
    Const FunctionName As String = "Escape"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Rx As VBScript_RegExp_55.RegExp
    Set Rx = CreateObject("VBScript.RegExp")
    Rx.Global = True
    Rx.MultiLine = True
    Rx.Pattern = "([""\\\/])"
    
    Dim Data As String
    Data = Rx.Replace(Value, "$1")              '// Escape double quotes, anti-slash and slah
    Data = Replace(Data, vbLf, "\n")            '// Escape new line
    Data = Replace(Data, vbCr, "\r")            '// Escape carriage return
    Data = Replace(Data, vbBack, "\b")          '// Escape backspace
    Data = Replace(Data, vbTab, "\t")           '// Escape tab
    Data = Replace(Data, vbVerticalTab, "\f")   '// Escape Vertical tab
    Data = EscapeUnicode(Data)                  '// Escape unicode characters
    Escape = Data
End Function

Private Function EscapeUnicode(ByVal Value As String) As String
    Const FunctionName As String = "EscapeUnicode"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Data As String
    Dim i As Long
    For i = 1 To Len(Value)
        Dim Character As String
        Character = Mid$(Value, i, 1)
        
        Select Case AscW(Character)
        Case 32 To 126
            Data = Data & Character
        Case Else
            Data = Data & "\u" & Right$("0000" & Hex$(AscW(Character)), 4)
        End Select
    Next
    EscapeUnicode = Data
End Function

Private Function Unescape(ByVal Value As String) As String
    Const FunctionName As String = "Unescape"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Data As String
    Data = UnescapeUnicode(Value)               '// Unescape unicode characters
    
    Dim Rx As VBScript_RegExp_55.RegExp
    Set Rx = CreateObject("VbScript.RegExp")
    Rx.Global = True
    Rx.MultiLine = True
    Rx.Pattern = "\\([""\\\/])"
    
    Data = Rx.Replace(Data, "$1")               '// Unescape double quotes, anti_slash and slash
    Data = Replace(Data, "\b", vbBack)          '// Unescape backspace
    Data = Replace(Data, "\f", vbVerticalTab)   '// Unescape vertical tab
    Data = Replace(Data, "\n", vbLf)            '// Unescape new line
    Data = Replace(Data, "\r", vbCr)            '// Unescape carriage return
    Data = Replace(Data, "\t", vbTab)           '// Unescape tab
    Unescape = Data
End Function

Private Function UnescapeUnicode(ByVal Value As String) As String
    Const FunctionName As String = "UnescapeUnicode"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Rx As VBScript_RegExp_55.RegExp
    Set Rx = CreateObject("VBScript.RegExp")
    Rx.Global = True
    Rx.MultiLine = True
    Rx.Pattern = "\\u([a-fA-F0-9]{4})"
    
    Dim Matchs As VBScript_RegExp_55.MatchCollection
    Dim Data As String
    Data = Value
    
    Dim Match As VBScript_RegExp_55.Match
    Set Matchs = Rx.Execute(Data)

    For Each Match In Matchs
        Rx.Pattern = "(\" & Match.Value & ")"
        Data = Rx.Replace(Data, ChrW$(Val("&H" & Match.SubMatches(0))))
    Next
    UnescapeUnicode = Data
End Function
