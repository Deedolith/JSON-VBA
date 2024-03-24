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
'@Folder "JSON"
Option Explicit

Private Const ModuleName As String = "JNumber"

Private mData As StringStream
Private mNumber As Double

'@Description "Constructor"
Friend Sub Create(ByVal StringStream As StringStream)
Attribute Create.VB_Description = "Constructor"
    Debug.Assert Not (StringStream Is Nothing)
    
    Const FunctionName As String = "Create"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Set mData = StringStream
    Parse
End Sub

'@Description "Class's data type"
Public Function DataType() As JType
Attribute DataType.VB_Description = "Class's data type"
    Const FunctionName As String = "DataType"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    DataType = JType.JSNumber
End Function

'@Description "Stream parsing"
Private Sub Parse()
Attribute Parse.VB_Description = "Stream parsing"
    Const FunctionName As String = "Parse"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim Value As String
    Value = mData.GetStringFromRegEx("^(?:-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?)")
    If (Value <> vbNullString) Then
        mData.EatString Value
    Else
        Err.Raise JException.JUnexpectedToken, "JNumber.Parse", "Numeric character expected."
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
    Const FunctionName As String = "Value (Get)"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Value = mNumber
End Property

'@Description "Set the numeric value of the object"
Public Property Let Value(ByVal Data As Double)
Attribute Value.VB_Description = "Set the numeric value of the object"
    Const FunctionName As String = "Value (Let)"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    mNumber = Data
End Property

'@Description "return a JSON string representation of the object"
Public Function ToJSONString() As String
Attribute ToJSONString.VB_Description = "return a JSON string representation of the object"
    Const FunctionName As String = "ToJSONString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim Output As String
    Output = ToString
    Output = Replace(Output, ",", ".")
    ToJSONString = Output
End Function

'@Description "Retrun a human readable string representation of the object"
Public Function ToString(Optional ByVal IndentMultiplier As Long = 0) As String
Attribute ToString.VB_Description = "Retrun a human readable string representation of the object"
    Debug.Assert IndentMultiplier >= 0
    Const FunctionName As String = "ToString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim Output As String
    Output = CStr(mNumber)
    ToString = Output
End Function
