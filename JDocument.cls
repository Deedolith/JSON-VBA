VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder("JSON")
Option Explicit
Implements JSObject

Private Const ModuleName As String = "JDocument"

Private mData As StringStream
Private mValue As JSObject

'@Description "Acquire data from a source"
Public Sub LoadFrom(ByVal Reader As IReader)
Attribute LoadFrom.VB_Description = "Acquire data from a source"
    Const FunctionName As String = "LoadFrom"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Set mData = Services.CreateStringStream(Reader.Execute)
    Dim Length As Long
    Length = Len(mData.Value)
On Error GoTo Error
    Parse
Exit Sub
Error:
    Err.Raise Err.Number, Err.Source, Err.Description & " at position " & Length - Len(mData.Value) & " (" & Left$(mData.Value, 20) & ") ."
End Sub

'@Description "Write Data to a target"
Public Sub WriteTo(ByVal Writer As IWriter)
Attribute WriteTo.VB_Description = "Write Data to a target"
    Const FunctionName As String = "WriteTo"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Writer.Execute mValue.ToJSONString
End Sub

'@Description "Retrieve the document's root JSON JSObject"
Public Function GetValueAs(ByVal DataType As JType) As JSObject
Attribute GetValueAs.VB_Description = "Retrieve the document's root JSON JSObject"
    Const FunctionName As String = "GetValueAs"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim JObject As JSObject
    Set JObject = Services.GetValueAs(mValue, DataType)
    Set GetValueAs = JObject
End Function

'@Description "Stream parsing"
Private Sub Parse()
Attribute Parse.VB_Description = "Stream parsing"
    Const FunctionName As String = "Parse"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Set mValue = Services.ParseValue(mData)
End Sub

'@Description "Retrieve a JSON element by querying the document with a pseudo XPath (case sensitive)."
Public Function Query(ByVal XPath As String) As JSObject
Attribute Query.VB_Description = "Retrieve a JSON element by querying the document with a pseudo XPath (case sensitive)."
    Const FunctionName As String = "Query"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim Path As Collection
    Set Path = ParseXpath(XPath)
    
    Dim JValue As JSObject
    Set JValue = mValue
    
    Dim Item As Variant
    For Each Item In Path
        If (Not IsNumeric(Item)) Then
            Dim JObject As JObject
            Set JObject = Services.GetValueAs(JValue, JType.JSObject)
            Set JValue = JObject.Members.Item(Item).Value
        Else
            Dim JArray As JArray
            Set JArray = Services.GetValueAs(JValue, JType.JSArray)
            Set JValue = JArray.Item(Item)
        End If
    Next
    Set Query = JValue
End Function

Private Function ParseXpath(ByVal XPath As String) As Collection
    Const FunctionName As String = "ParseXpath"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim Output As Collection
    Set Output = New Collection

    Dim StringStream As StringStream
    Set StringStream = Services.CreateStringStream(XPath)
    
    StringStream.EatCharacter "/"
    
    Dim Character As String
    Do While Not StringStream.EOF
        Character = StringStream.PeekCharacter
        If (Character = "[") Then
            StringStream.EatCharacter "["
            
            Dim Number As String
            Number = ParseNumber(StringStream)
            Output.Add CInt(Number)
            
            StringStream.EatCharacter "]"
        ElseIf (Character <> "/") Then
            Dim StringData As String
            StringData = ParseString(StringStream)
            Output.Add StringData
        Else
            StringStream.EatCharacter "/"
        End If
    Loop
    Set ParseXpath = Output
End Function

Private Function ParseNumber(ByVal StringStream As StringStream) As String
    Debug.Assert Not (StringStream Is Nothing)
    
    Const FunctionName As String = "ParseNumber"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim Number As String
    Number = StringStream.GetStringFromRegEx("^[1-9]\d*")
    StringStream.EatString Number
    ParseNumber = Number
End Function

Private Function ParseString(ByVal StringStream As StringStream) As String
    Debug.Assert Not (StringStream Is Nothing)
    
    Const FunctionName As String = "ParseString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim Data As String
    Data = StringStream.GetStringFromRegEx("[^\[\/]+")
    StringStream.EatString Data
    ParseString = Data
End Function

Public Function ToString() As String
    ToString = mValue.ToString
End Function

Private Function JSObject_ToJSONString() As String
    Const FunctionName As String = "JSObject_ToJSONString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    JSObject_ToJSONString = ToString
End Function

Private Function JSObject_ToString(Optional ByVal IndentMultiplier As Long = 0&) As String
    Const FunctionName As String = "JSObject_ToString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    JSObject_ToString = ToString()
End Function
