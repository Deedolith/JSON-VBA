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

Private mData As String
Private Const ModuleName As String = "StringStream"

'@Description "Constructor"
Friend Sub Create(ByVal Data As String)
Attribute Create.VB_Description = "Constructor"
#If DEV Then
    Const FunctionName As String = "Create"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    mData = Data
End Sub

'@Description "Eat any possible space"
Public Sub DiscardSpaces()
Attribute DiscardSpaces.VB_Description = "Eat any possible space"
#If DEV Then
    Const FunctionName As String = "DiscardSpaces"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim rx As Object        '// VBScript_RegExp_55.RegExp
    Set rx = VBA.CreateObject("VBScript.RegExp")

    rx.Pattern = "\s*([\s\S]*)"
    mData = rx.Replace(mData, "$1")
End Sub

'@Description "Eat a single character"
Public Sub EatCharacter(ByVal Character As String)
Attribute EatCharacter.VB_Description = "Eat a single character"
    Debug.Assert (Len(Character) = 1)   '// Expected single character

#If DEV Then
    Const FunctionName As String = "EatCharacter"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    If (Left$(mData, 1) = Character) Then
        mData = Right$(mData, Len(mData) - 1)
    Else
        Err.Raise JSON.JException.JUnexpectedCharacter, "StringStream.EatCharacter", """" & Character & """ Expected"
    End If
End Sub

'@Description "Lookup at the next character in the stream"
Public Function PeekCharacter() As String
Attribute PeekCharacter.VB_Description = "Lookup at the next character in the stream"
#If DEV Then
    Const FunctionName As String = "PeekCharacter"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    PeekCharacter = Left$(mData, 1)
End Function

'@DefaultMember
'@Description "Return the current stream value"
Public Property Get Value() As String
Attribute Value.VB_Description = "Return the current stream value"
Attribute Value.VB_UserMemId = 0
#If DEV Then
    Const FunctionName As String = "Value"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Value = mData
End Property

'@Description "Extract the first occurence matching the regex's first group"
Public Function GetStringFromRegEx(ByVal RegEx As String) As String
Attribute GetStringFromRegEx.VB_Description = "Extract the first occurence matching the regex's first group"
#If DEV Then
    Const FunctionName As String = "GetStringFromRegEx"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim rx As Object        '// VBScript_RegExp_55.RegExp
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Pattern = RegEx
    rx.Global = True
    rx.MultiLine = True
    
    Dim Matchs As Object        '// VBScript_RegExp_55.MatchCollection
    Set Matchs = rx.Execute(mData)
    If Not (Matchs Is Nothing) Then
        If (Matchs.Count) Then
            GetStringFromRegEx = Matchs.Item(0).Value
        End If
    End If
End Function

'@Description "Eat a string"
Public Sub EatString(ByVal Data As String)
Attribute EatString.VB_Description = "Eat a string"
#If DEV Then
    Const FunctionName As String = "EatString"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    If (InStr(mData, Data) = 1) Then
        mData = Right$(mData, Len(mData) - Len(Data))
    Else
        Err.Raise JSON.JException.JUnexpectedCharacter, "StringStream.EatString", "String Expected"
    End If
End Sub

'@Description "End of Stream indicator"
Public Function EOF() As Boolean
Attribute EOF.VB_Description = "End of Stream indicator"
#If DEV Then
    Const FunctionName As String = "EOF"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    EOF = mData = vbNullString
End Function
