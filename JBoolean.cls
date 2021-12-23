VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JBoolean"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder "JSON"
Option Explicit

'@TODO Implement the JBoolean class
Const ModuleName As String = "JBoolean"
Private mValue As Boolean
Private mData As JSON.StringStream

Friend Sub Create(ByRef SS As JSON.StringStream)
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
Public Property Get DataType() As JSON.JType
Attribute DataType.VB_Description = "Class's data type."
#If DEV Then
    Const FunctionName As String = "DataType"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    DataType = JSON.JType.JSBoolean
End Property

'@Description "Return a human readable string representation of the class."
Public Function ToString(Optional ByVal IndentMultiplier As Long = 0) As String
Attribute ToString.VB_Description = "Return a human readable string representation of the class."
#If DEV Then
    Const FunctionName As String = "ToString"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Output As String
    If (mValue) Then
        Output = "true"
    Else
        Output = "false"
    End If
    ToString = Output
End Function

'@Description "Return string representation of the class."
Public Function ToJSONString() As String
Attribute ToJSONString.VB_Description = "Return string representation of the class."
#If DEV Then
    Const FunctionName As String = "ToJSONString"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Output As String
    Output = ToString
    ToJSONString = Output
End Function

'@Description "true or false"
Public Property Get Value() As Boolean
Attribute Value.VB_Description = "true or false"
#If DEV Then
    Const FunctionName As String = "Value (Get)"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Value = mValue
End Property

'@Description "true or false"
Public Property Let Value(ByVal Data As Boolean)
Attribute Value.VB_Description = "true or false"
#If DEV Then
    Const FunctionName As String = "Value (Let)"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    mValue = Data
End Property

'@TODO Implement the Parse member
Private Sub Parse()
#If DEV Then
    Const FunctionName As String = "Parse"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Value As String
    Value = mData.GetStringFromRegEx("^(?:true|false)")
    If ((Value = "true") Or (Value = "false")) Then
        mData.EatString Value
        mValue = Value = "true"
    Else
        Err.Raise JSON.JSException.JSUnexpectedToken, "JBoolean.Parse", "Expecting true or false."
    End If
End Sub
