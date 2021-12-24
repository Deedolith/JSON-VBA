VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JNull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder("JSON")
Option Explicit

Private Const ModuleName As String = "JNull"
Private mData As JSON.StringStream

'@Description "Constructor"
Friend Sub Create(ByRef SS As JSON.StringStream)
Attribute Create.VB_Description = "Constructor"
    Debug.Assert Not (SS Is Nothing)

#If DEV Then
    Const FunctionName As String = "Create"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Set mData = SS
    Parse
End Sub

'@Description "Class's data type."
Public Function DataType() As JSON.JType
Attribute DataType.VB_Description = "Class's data type."
#If DEV Then
    Const FunctionName As String = "DataType"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    DataType = JSON.JType.JSNull
End Function

'@Description "Stream parsing"
Private Sub Parse()
Attribute Parse.VB_Description = "Stream parsing"
#If DEV Then
    Const FunctionName As String = "Parse"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Value As String
    Value = mData.GetStringFromRegEx("^null")
    If (Value = "null") Then
        mData.EatString Value
    Else
        Err.Raise JSON.JException.JUnexpectedToken, "JNull.Parse", """null"" expected."
    End If
End Sub

'@Description return a Null value
Public Property Get Value() As Variant
#If DEV Then
    Const FunctionName As String = "Value (Get)"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Value = Null
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
    Output = ToString
    ToJSONString = Output
End Function

'@Description "Return a human readable string representation of the object"
Public Function ToString(Optional ByVal IndentMultiplier As Long = 0) As String
Attribute ToString.VB_Description = "Return a human readable string representation of the object"
    Debug.Assert IndentMultiplier >= 0

#If DEV Then
    Const FunctionName As String = "ToString"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Output As String
    Output = "null"
    ToString = Output
End Function
