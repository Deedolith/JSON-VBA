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
'@Folder "JSON"
Option Explicit

Private Const ModuleName As String = "JNull"

Private mData As StringStream

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

'@Description "Class's data type."
Public Function DataType() As JType
Attribute DataType.VB_Description = "Class's data type."
    Const FunctionName As String = "DataType"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    DataType = JType.JSNull
End Function

'@Description "Stream parsing"
Private Sub Parse()
Attribute Parse.VB_Description = "Stream parsing"
    Const FunctionName As String = "Parse"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim Value As String
    Value = mData.GetStringFromRegEx("^null")
    If (Value = "null") Then
        mData.EatString Value
    Else
        Err.Raise JException.JUnexpectedToken, "JNull.Parse", """null"" expected."
    End If
End Sub

'@Description return a Null value
Public Property Get Value() As Variant
    Const FunctionName As String = "Value (Get)"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Value = Null
End Property

'@Description "Return a JSON string representation of the object"
Public Function ToJSONString() As String
Attribute ToJSONString.VB_Description = "Return a JSON string representation of the object"
    Const FunctionName As String = "ToJSONString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim Output As String
    Output = ToString
    ToJSONString = Output
End Function

'@Description "Return a human readable string representation of the object"
Public Function ToString(Optional ByVal IndentMultiplier As Long = 0) As String
Attribute ToString.VB_Description = "Return a human readable string representation of the object"
    Debug.Assert IndentMultiplier >= 0
    
    Const FunctionName As String = "ToString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim Output As String
    Output = "null"
    ToString = Output
End Function
