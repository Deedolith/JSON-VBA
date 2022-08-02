VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represent a JSON object."
'@Exposed
'@ModuleDescription "Represent a JSON object."
'@IgnoreModule AssignmentNotUsed, VariableNotUsed
'@Folder("JSON")
Option Explicit

Private mData As JSON.StringStream
Private mMembers As JSON.Members
#If DEV Then
    Private Const ModuleName As String = "JObject"
#End If

'@Description "Constructor"
Friend Sub Create(ByVal StringStream As JSON.StringStream)
Attribute Create.VB_Description = "Constructor"
    Debug.Assert Not (StringStream Is Nothing)      '// Data must be a valid instance
    
#If DEV Then
    Const FunctionName As String = "Create"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Set mData = StringStream
    Parse
End Sub

Private Sub Class_Initialize()
#If DEV Then
    Const FunctionName As String = "Class_Initialize"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Set mMembers = New JSON.Members
End Sub

'@Description "Stream parsing"
Private Sub Parse()
Attribute Parse.VB_Description = "Stream parsing"
#If DEV Then
    Const FunctionName As String = "Parse"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

 mData.EatCharacter "{"
    Do
        Dim DoLoop As Boolean
        DoLoop = False

        mData.DiscardSpaces
        If (mData.PeekCharacter <> "}") Then
            ParseMembers mData
            mData.DiscardSpaces
            If (mData.PeekCharacter = ",") Then
                mData.EatCharacter (",")
                DoLoop = True
            End If
        End If
    Loop While DoLoop
    mData.EatCharacter "}"
End Sub

'@Description "Parse elements within the object"
Private Sub ParseMembers(ByVal StringStream As JSON.StringStream)
Attribute ParseMembers.VB_Description = "Parse elements within the object"
    Debug.Assert Not StringStream Is Nothing

#If DEV Then
    Const FunctionName As String = "ParseMembers"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    StringStream.EatCharacter DOUBLEQUOTE
    
    Dim Name As String
    Name = StringStream.GetStringFromRegEx("^(?:\\(?:[""\\\/bfnrt]|u[a-fA-F0-9]{4})|[^""\\\0-\x1F\x7F]+)*")
    StringStream.EatString Name
    Name = Services.Unescape(Name)
    StringStream.EatCharacter DOUBLEQUOTE
    StringStream.DiscardSpaces
    StringStream.EatCharacter SEMICOLON
    StringStream.DiscardSpaces
    
    Dim Value As Object
    Set Value = Services.ParseValue(StringStream)
    
    Dim Pair As JSON.Pair
    Set Pair = Services.CreatePair(Name, Value)
    mMembers.Add Pair
End Sub

'@Description "Class's data type."
Public Function DataType() As JSON.JType
Attribute DataType.VB_Description = "Class's data type."
#If DEV Then
    Const FunctionName As String = "DataType"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    DataType = JSON.JType.JSObject
End Function

'@Description "The object's default value"
Public Property Get Value() As String
Attribute Value.VB_Description = "The object's default value"
#If DEV Then
    Const FunctionName As String = "Value"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Value = "[Object]"
End Property

'@Description "Return a human readable string representation of the object"
Public Function ToString(Optional ByVal IndentMultiplier As Long = 0) As String
Attribute ToString.VB_Description = "Return a human readable string representation of the object"
#If DEV Then
    Const FunctionName As String = "ToString"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If
    
    Dim Tabs As String
    Dim i As Long
    For i = 0 To IndentMultiplier - 1
        Tabs = Tabs & vbTab
    Next

    Dim Output As String
    Output = Output & "{" & vbCrLf

    Dim SecondLoop As Boolean
    SecondLoop = False
    
    Dim Pair As JSON.Pair
    For Each Pair In mMembers
        If (SecondLoop) Then
            Output = Output & "," & vbCrLf
        End If
        Output = Output & Tabs & vbTab & DOUBLEQUOTE & Pair.Name & """:" & Pair.Value.ToString(IndentMultiplier + 1)
        SecondLoop = True
    Next
    Output = Output & vbCrLf
    Output = Output & Tabs & "}"
    ToString = Output
End Function

'@Description "Return a JSON string representation of the object"
Public Function ToJSONString() As String
Attribute ToJSONString.VB_Description = "Return a JSON string representation of the object"
#If DEV Then
    Const FunctionName As String = "ToJSONString"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Output As String
    Output = Output & "{"

    Dim SecondLoop As Boolean
    SecondLoop = False
    
    Dim Pair As JSON.Pair
    For Each Pair In mMembers
        If (SecondLoop) Then
            Output = Output & ","
        End If
        Output = Output & """" & Services.Escape(Pair.Name) & """:" & Pair.Value.ToJSONString
        SecondLoop = True
    Next
    Output = Output & "}"
    ToJSONString = Output
End Function

Public Property Get Members() As JSON.Members
#If DEV Then
    Const FunctionName As String = "Members"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Set Members = mMembers
End Property
