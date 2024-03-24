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
Implements JSObject

Private Const ModuleName As String = "JString"

Private mData As StringStream
Private mValue As String

'@Description "Constructor"
Friend Sub Create(ByVal StringStream As StringStream)
Attribute Create.VB_Description = "Constructor"
    Debug.Assert Not StringStream Is Nothing

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

    DataType = JType.JSString
End Property

'@Description "Stream parsing"
Private Sub Parse()
Attribute Parse.VB_Description = "Stream parsing"
    Const FunctionName As String = "Parse"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    mData.EatCharacter DOUBLEQUOTE
    Dim Value As String
    Value = mData.GetStringFromRegEx("^(?:\\(?:[""\\\/bfnrt]|u[a-fA-F0-9]{4})|[^""\\\0-\x1F\x7F]+)*")
    mData.EatString Value
    mData.EatCharacter DOUBLEQUOTE
    Value = Services.Unescape(Value)
    mValue = Value
End Sub

'@Description "Return the string value of the JSObject"
Public Property Get Value() As String
Attribute Value.VB_Description = "Return the string value of the JSObject"
    Const FunctionName As String = "Value (Get)"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Value = Services.Escape(mValue)
End Property

'@Description "Set the string value of the JSObject"
Public Property Let Value(ByVal Data As String)
Attribute Value.VB_Description = "Set the string value of the JSObject"
    Const FunctionName As String = "Value (Let)"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    mValue = Data
End Property

'@Description "Return a JSON string representation of the JSObject"
Public Function ToJSONString() As String
Attribute ToJSONString.VB_Description = "Return a JSON string representation of the JSObject"
    Const FunctionName As String = "ToJSONString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim Output As String
    Output = Output & DOUBLEQUOTE
    Output = Output & Services.Escape(mValue)
    Output = Output & DOUBLEQUOTE
    ToJSONString = Output
End Function

'@Description "Retrun a human readable string representation of the JSObject"
Public Function ToString(Optional ByVal IndentMultiplier As Long = 0) As String
Attribute ToString.VB_Description = "Retrun a human readable string representation of the JSObject"
    Debug.Assert IndentMultiplier >= 0
    Const FunctionName As String = "ToString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim Output As String
    Output = Output & DOUBLEQUOTE
    Output = Output & mValue
    Output = Output & DOUBLEQUOTE
    ToString = Output
End Function

Private Function JSObject_ToJSONString() As String
    Const FunctionName As String = "JSObject_ToJSONString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    JSObject_ToJSONString = ToJSONString
End Function

Private Function JSObject_ToString(Optional ByVal IndentMultiplier As Long = 0&) As String
    Const FunctionName As String = "JSObject_ToString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    JSObject_ToString = ToString(IndentMultiplier)
End Function
