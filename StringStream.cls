VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule AssignmentNotUsed, VariableNotUsed
'@Folder "JSON"
Option Explicit

Private Const ModuleName As String = "StringStream"

Private mData As String

'@Description "Constructor"
Friend Sub Create(ByVal Data As String)
Attribute Create.VB_Description = "Constructor"
    Const FunctionName As String = "Create"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    mData = Data
End Sub

'@Description "Eat any possible space"
Public Sub DiscardSpaces()
Attribute DiscardSpaces.VB_Description = "Eat any possible space"
    Const FunctionName As String = "DiscardSpaces"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim RegEx As Object        '// VBScript_RegExp_55.RegExp
    Set RegEx = VBA.CreateObject("VBScript.RegExp")

    RegEx.Pattern = "\s*([\s\S]*)"
    mData = RegEx.Replace(mData, "$1")
End Sub

'@Description "Eat a single character"
Public Sub EatCharacter(ByVal Character As String)
Attribute EatCharacter.VB_Description = "Eat a single character"
    Debug.Assert (Len(Character) = 1)   '// Expected single character
    
    Const FunctionName As String = "EatCharacter"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    If (Left$(mData, 1) = Character) Then
        mData = Right$(mData, Len(mData) - 1)
    Else
        Err.Raise JException.JUnexpectedCharacter, "StringStream.EatCharacter", """" & Character & """ Expected"
    End If
End Sub

'@Description "Lookup at the next character in the stream"
Public Function PeekCharacter() As String
Attribute PeekCharacter.VB_Description = "Lookup at the next character in the stream"
    Const FunctionName As String = "PeekCharacter"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    PeekCharacter = Left$(mData, 1)
End Function

'@DefaultMember
'@Description "Return the current stream value"
Public Property Get Value() As String
Attribute Value.VB_Description = "Return the current stream value"
Attribute Value.VB_UserMemId = 0
    Const FunctionName As String = "Value"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Value = mData
End Property

'@Description "Extract the first occurence matching the regex's first group"
Public Function GetStringFromRegEx(ByVal Pattern As String) As String
Attribute GetStringFromRegEx.VB_Description = "Extract the first occurence matching the regex's first group"
    Const FunctionName As String = "GetStringFromRegEx"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim RegEx As Object        '// VBScript_RegExp_55.RegExp
    Set RegEx = VBA.CreateObject("VBScript.RegExp")
    RegEx.Pattern = Pattern
    
    Dim Matchs As Object        '// VBScript_RegExp_55.MatchCollection
    Set Matchs = RegEx.Execute(mData)
    If Not (Matchs Is Nothing) Then
        If (Matchs.Count) Then
            GetStringFromRegEx = Matchs.Item(0).Value
        End If
    End If
End Function

'@Description "Eat a string"
Public Sub EatString(ByVal Data As String)
Attribute EatString.VB_Description = "Eat a string"
    Const FunctionName As String = "EatString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    If (InStr(mData, Data) = 1) Then
        mData = Right$(mData, Len(mData) - Len(Data))
    Else
        Err.Raise JException.JUnexpectedCharacter, "StringStream.EatString", "String Expected"
    End If
End Sub

'@Description "End of Stream indicator"
Public Function EOF() As Boolean
Attribute EOF.VB_Description = "End of Stream indicator"
    Const FunctionName As String = "EOF"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    EOF = mData = vbNullString
End Function
