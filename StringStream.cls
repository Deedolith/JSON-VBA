VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder("JSON.DataReaders")
Option Explicit

Private Const ModuleName As String = "StringStream"

Private Type TData
    Value As String
End Type

Private This As TData

'@Description "Constructor"
Friend Sub Create(ByVal Data As String)
Attribute Create.VB_Description = "Constructor"
    Const FunctionName As String = "Create"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    This.Value = Data
End Sub

'@Description "Eat any possible spaces"
Public Sub EatSpaces()
Attribute EatSpaces.VB_Description = "Eat any possible spaces"
    Const FunctionName As String = "EatSpaces"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Rx As VBScript_RegExp_55.RegExp
    Set Rx = CreateObject("VBScript.RegExp")
    Rx.Pattern = "^[\s]*([\s\S]+)"
    This.Value = Rx.Replace(This.Value, "$1")
End Sub

'@Description "Eat a single character"
Public Sub EatCharacter(ByVal Character As String)
Attribute EatCharacter.VB_Description = "Eat a single character"
    Const FunctionName As String = "EatCharacter"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Debug.Assert Len(Character) = 1     '// Expected single character
    
    If (PeekCharacter = Character) Then
        This.Value = Right$(This.Value, Len(This.Value) - 1)
    Else
        Err.Raise JsonExceptionUnexpectedCharacter, ModuleName & "." & FunctionName, "Character """ & Character & """ excpected."
    End If
End Sub

'@Description "Lookup the next character in the stream"
Public Function PeekCharacter() As String
Attribute PeekCharacter.VB_Description = "Lookup the next character in the stream"
    Const FunctionName As String = "PeekCharacter"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    PeekCharacter = Left$(This.Value, 1)
End Function

'@Description "Get the stream's current value"
Public Function Value() As String
Attribute Value.VB_Description = "Get the stream's current value"
    Const FunctionName As String = "Value"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Value = This.Value
End Function

'@Description "End of stream indicator"
Public Function EOF() As Boolean
Attribute EOF.VB_Description = "End of stream indicator"
    Const FunctionName As String = "EOF"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    EOF = This.Value = vbNullString
End Function

'@Description "Check wether the stream's value match the given regular expression"
Public Function Match(ByVal RegEx As String) As Boolean
Attribute Match.VB_Description = "Check wether the stream's value match the given regular expression"
    Const FunctionName As String = "Match"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Rx As VBScript_RegExp_55.RegExp
    Set Rx = CreateObject("VBScript.RegExp")
    Rx.Global = True
    Rx.IgnoreCase = False
    Rx.Pattern = RegEx
    Match = Rx.test(This.Value)
End Function

'@Description "Eat a string"
Public Sub EatString(ByVal Data As String)
Attribute EatString.VB_Description = "Eat a string"
    Const FunctionName As String = "EatString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    If (InStr(This.Value, Data) = 1) Then
        This.Value = Right$(This.Value, Len(This.Value) - Len(Data))
    Else
        Err.Raise JsonExceptionUnexpectedCharacter, ModuleName & "." & FunctionName, "String like """ & Left$(This.Value, 20) & """ expected."
    End If
End Sub

'@Description "Return the 1st value matching the given regular expression"
Public Function PeekString(ByVal RegEx As String) As String
Attribute PeekString.VB_Description = "Return the 1st value matching the given regular expression"
    Const FunctionName As String = "EatString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Rx As VBScript_RegExp_55.RegExp
    Set Rx = CreateObject("VBScript.RegExp")
    Rx.Global = True
    Rx.IgnoreCase = False
    Rx.Pattern = RegEx
    
    Dim Matchs As VBScript_RegExp_55.MatchCollection
    Set Matchs = Rx.Execute(This.Value)
    
    Debug.Assert Matchs.Count > 0    '// Provided regex did not match
    
    PeekString = Matchs.Item(0).Value
End Function
