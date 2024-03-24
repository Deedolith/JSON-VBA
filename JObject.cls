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

Private Const ModuleName As String = "JObject"

Private mData As StringStream
Private mMembers As Members

'@Description "Constructor"
Friend Sub Create(ByVal StringStream As StringStream)
Attribute Create.VB_Description = "Constructor"
    Debug.Assert Not (StringStream Is Nothing)      '// Data must be a valid instance
    
    Const FunctionName As String = "Create"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Set mData = StringStream
    Parse
End Sub

Private Sub Class_Initialize()
    Const FunctionName As String = "Class_Initialize"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Set mMembers = New Members
End Sub

'@Description "Stream parsing"
Private Sub Parse()
Attribute Parse.VB_Description = "Stream parsing"
    Const FunctionName As String = "Parse"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

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
Private Sub ParseMembers(ByVal StringStream As StringStream)
Attribute ParseMembers.VB_Description = "Parse elements within the object"
    Debug.Assert Not StringStream Is Nothing
    Const FunctionName As String = "ParseMembers"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

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
    
    Dim Pair As Pair
    Set Pair = Services.CreatePair(Name, Value)
    mMembers.Add Pair
End Sub

'@Description "Class's data type."
Public Function DataType() As JType
Attribute DataType.VB_Description = "Class's data type."
    Const FunctionName As String = "DataType"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    DataType = JType.JSobject
End Function

'@Description "The object's default value"
Public Property Get Value() As String
Attribute Value.VB_Description = "The object's default value"
    Const FunctionName As String = "Value"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Value = "[Object]"
End Property

'@Description "Return a human readable string representation of the object"
Public Function ToString(Optional ByVal IndentMultiplier As Long = 0) As String
Attribute ToString.VB_Description = "Return a human readable string representation of the object"
    Const FunctionName As String = "ToString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Tabs As String
    Dim i As Long
    For i = 0 To IndentMultiplier - 1
        Tabs = Tabs & vbTab
    Next

    Dim Output As String
    Output = Output & "{" & vbCrLf

    Dim SecondLoop As Boolean
    SecondLoop = False
    
    Dim Pair As Pair
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
    Const FunctionName As String = "ToJSONString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim Output As String
    Output = Output & "{"

    Dim SecondLoop As Boolean
    SecondLoop = False
    
    Dim Pair As Pair
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

Public Property Get Members() As Members
    Const FunctionName As String = "Members"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Set Members = mMembers
End Property
