VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JBoolean"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder "JSON"
Option Explicit

Const ModuleName As String = "JBoolean"

Private mValue As Boolean
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
Public Property Get DataType() As JType
Attribute DataType.VB_Description = "Class's data type."
    Const FunctionName As String = "DataType"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    DataType = JType.JSBoolean
End Property

'@Description "Return a human readable string representation of the object"
Public Function ToString(Optional ByVal IndentMultiplier As Long = 0) As String
Attribute ToString.VB_Description = "Return a human readable string representation of the object"
    Const FunctionName As String = "ToString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim Output As String
    If (mValue) Then
        Output = "true"
    Else
        Output = "false"
    End If
    ToString = Output
End Function

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

'@Description "true or false"
Public Property Get Value() As Boolean
Attribute Value.VB_Description = "true or false"
    Const FunctionName As String = "Value (Get)"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Value = mValue
End Property

'@Description "true or false"
Public Property Let Value(ByVal Data As Boolean)
Attribute Value.VB_Description = "true or false"
    Const FunctionName As String = "Value (Let)"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    mValue = Data
End Property

'@Description "Stream parsing"
Private Sub Parse()
Attribute Parse.VB_Description = "Stream parsing"
    Const FunctionName As String = "Parse"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim Value As String
    Value = mData.GetStringFromRegEx("^(?:true|false)")
    If ((Value = "true") Or (Value = "false")) Then
        mData.EatString Value
        mValue = Value = "true"
    Else
        Err.Raise JException.JUnexpectedToken, "JBoolean.Parse", """true"" or ""false"" expected."
    End If
End Sub
