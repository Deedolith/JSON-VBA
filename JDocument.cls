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

Private mData As JSON.StringStream
Private mValue As Object
#If DEV Then
    Private Const ModuleName As String = "JDocument"
#End If

'@Description "Acquire data from a source"
Public Sub LoadFrom(ByVal Reader As IReader)
Attribute LoadFrom.VB_Description = "Acquire data from a source"
#If DEV Then
    Const FunctionName As String = "LoadFrom"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

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
#If DEV Then
    Const FunctionName As String = "WriteTo"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Writer.Execute mValue.ToJSONString
End Sub

'@Description "Retrieve the document's root JSON object"
Public Function GetValueAs(ByVal DataType As JSON.JType) As Object
Attribute GetValueAs.VB_Description = "Retrieve the document's root JSON object"
#If DEV Then
    Const FunctionName As String = "GetValueAs"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim JObject As Object
    Set JObject = Services.GetValueAs(mValue, DataType)
    Set GetValueAs = JObject
End Function

'@Description "Stream parsing"
Private Sub Parse()
Attribute Parse.VB_Description = "Stream parsing"
#If DEV Then
    Const FunctionName As String = "Parse"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Set mValue = Services.ParseValue(mData)
End Sub

'@Description "Retrieve a JSON element by querying the document with a pseudo XPath (case sensitive)."
Public Function Query(ByVal XPath As String) As Object
Attribute Query.VB_Description = "Retrieve a JSON element by querying the document with a pseudo XPath (case sensitive)."
#If DEV Then
    Const FunctionName As String = "Query"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Path As Collection
    Set Path = ParseXpath(XPath)
    
    Dim JValue As Object
    Set JValue = mValue
    
    Dim Item As Variant
    For Each Item In Path
        If (Not IsNumeric(Item)) Then
            Dim JObject As JSON.JObject
            Set JObject = Services.GetValueAs(JValue, JSON.JType.JSObject)
            Set JValue = JObject.Members.Item(Item).Value
        Else
            Dim Jarray As JSON.Jarray
            Set Jarray = Services.GetValueAs(JValue, JSON.JType.JSArray)
            Set JValue = Jarray.Item(Item)
        End If
    Next
    Set Query = JValue
End Function

Private Function ParseXpath(ByVal XPath As String) As Collection
#If DEV Then
    Const FunctionName As String = "ParseXpath"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Output As Collection
    Set Output = New Collection

    Dim StringStream As JSON.StringStream
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

Private Function ParseNumber(ByVal StringStream As JSON.StringStream) As String
#If DEV Then
    Debug.Assert Not (StringStream Is Nothing)
    Const FunctionName As String = "ParseNumber"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Number As String
    Number = StringStream.GetStringFromRegEx("^[1-9]\d*")
    StringStream.EatString Number
    ParseNumber = Number
End Function

Private Function ParseString(ByVal StringStream As JSON.StringStream) As String
#If DEV Then
    Debug.Assert Not (StringStream Is Nothing)
    Const FunctionName As String = "ParseString"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Data As String
    Data = StringStream.GetStringFromRegEx("[^\[\/]+")
    StringStream.EatString Data
    ParseString = Data
End Function
