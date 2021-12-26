VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder "JSON"
Option Explicit

Private mData As JSON.StringStream
Private mValue As String
Private Const ModuleName As String = "JString"

'@Description "Constructor"
Friend Sub Create(ByRef SS As JSON.StringStream)
Attribute Create.VB_Description = "Constructor"
    Debug.Assert Not SS Is Nothing

#If DEV Then
    Const FunctionName As String = "Create"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Set mData = SS
    Parse
End Sub

'@Description "Class's data type."
Public Property Get DataType() As JSON.JType
Attribute DataType.VB_Description = "Class's data type."
#If DEV Then
    Const FunctionName As String = "DataType"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    DataType = JSON.JType.JSString
End Property

'@Description "Stream parsing"
Private Sub Parse()
Attribute Parse.VB_Description = "Stream parsing"
#If DEV Then
    Const FunctionName As String = "Parse"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    mData.EatCharacter """"
    Dim Value As String
    Value = mData.GetStringFromRegEx("^(?:\\(?:[""\\\/bfnrt]|u[a-fA-F0-9]{4})|[^""\\\0-\x1F\x7F]+)*")
    mData.EatString Value
    mData.EatCharacter """"
    Value = Services.Unescape(Value)
    mValue = Value
End Sub

'@Description "Return the string value of the object"
Public Property Get Value() As String
Attribute Value.VB_Description = "Return the string value of the object"
#If DEV Then
    Const FunctionName As String = "Value (Get)"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Value = Services.Escape(mValue)
End Property

'@Description "Set the string value of the object"
Public Property Let Value(ByVal Data As String)
Attribute Value.VB_Description = "Set the string value of the object"
#If DEV Then
    Const FunctionName As String = "Value (Let)"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    mValue = Data
End Property

'@Description "Return a JSON string representation of the object"
Public Function ToJSONString() As String
Attribute ToJSONString.VB_Description = "Return a JSON string representation of the object"
#If DEV Then
    Const FunctionName As String = "ToJSONString"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Output As String
    Output = Output & """"
    Output = Output & Services.Escape(mValue)
    Output = Output & """"
    ToJSONString = Output
End Function

'@Description "Retrun a human readable string representation of the object"
Public Function ToString(Optional ByVal IndentMultiplier As Long = 0) As String
Attribute ToString.VB_Description = "Retrun a human readable string representation of the object"
    Debug.Assert IndentMultiplier >= 0

#If DEV Then
    Const FunctionName As String = "ToString"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Output As String
    Output = Output & """"
    Output = Output & mValue
    Output = Output & """"
    ToString = Output
End Function
