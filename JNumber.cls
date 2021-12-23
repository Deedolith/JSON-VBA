VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder("JSON")
Option Explicit

Private mData As JSON.StringStream
Private mNumber As Double
Private Const ModuleName As String = "JNumber"

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

Public Function DataType() As JSON.JType
#If DEV Then
    Const FunctionName As String = "DataType"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    DataType = JSON.JType.JSNumber
End Function

Private Sub Parse()
#If DEV Then
    Const FunctionName As String = "Parse"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Value As String
    Value = mData.GetStringFromRegEx("^(?:-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?)")
    If (Value <> vbNullString) Then
        mData.EatString Value
    Else
        Err.Raise JSON.JSException.JSUnexpectedToken, "JNumber.Parse", "Numeric character expected."
    End If
On Error GoTo RegionalError
    mNumber = CDbl(Value)
Exit Sub
regionalResume:
    mNumber = CDbl(Value)
Exit Sub
RegionalError:
    If (Err.Number = 13) Then   '// Type mismatch due to regional settings
        Value = Replace(Value, ".", ",")
    End If
    Err.Clear
    GoTo regionalResume
End Sub

'@Description "Return the numeric value of the object."
Public Property Get Value() As Double
Attribute Value.VB_Description = "Return the numeric value of the object."
#If DEV Then
    Const FunctionName As String = "Value (Get)"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Value = mNumber
End Property

Public Property Let Value(ByVal Data As Double)
#If DEV Then
    Const FunctionName As String = "Value (Let)"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    mNumber = Data
End Property

Public Function ToJSONString() As String
#If DEV Then
    Const FunctionName As String = "ToJSONString"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Output As String
    Output = ToString
    Output = Replace(Output, ",", ".")
    ToJSONString = Output
End Function

Public Function ToString(Optional ByVal IndentMultiplier As Long = 0) As String
    Debug.Assert IndentMultiplier >= 0

#If DEV Then
    Const FunctionName As String = "ToString"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Output As String
    Output = CStr(mNumber)
    ToString = Output
End Function
