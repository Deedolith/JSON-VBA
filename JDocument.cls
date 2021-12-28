VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("JSON")
Option Explicit

Private mData As JSON.StringStream
Private mValue As Object
Private Const ModuleName As String = "JDocument"

'@Description "Acquire data from a source"
Public Sub LoadFrom(ByRef Reader As IReader)
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
    Err.Raise Err.Number, Err.Source, Err.Description & " at position " & Length - Len(mData.Value) & " (" & Left(mData.Value, 20) & ") ."
End Sub

'@Description "Write Data to a target"
Public Sub WriteTo(ByRef Writer As IWriter)
#If DEV Then
    Const FunctionName As String = "WriteTo"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Writer.Execute mValue.ToJSONString
End Sub

'@Description "Retrieve the document's root JSON object"
Public Function GetValueAs(ByVal DataType As JSON.JType) As Object
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
Private Function Parse()
#If DEV Then
    Const FunctionName As String = "Parse"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If
    Set mValue = Services.ParseValue(mData)
End Function

'@Description "Retrieve a JSON element by querying the document with a pseudo XPath (case sensitive)."
Public Function Query(ByVal XPath As String) As Object
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
            Set JValue = JObject.Members(Item).Value
        Else
            Dim JArray As JSON.JArray
            Set JArray = Services.GetValueAs(JValue, JSON.JType.JSArray)
            Set JValue = JArray(Item)
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

    Dim SS As StringStream
    Set SS = Services.CreateStringStream(XPath)
    
    SS.EatCharacter "/"
    
    Dim Character As String
    Do While Not SS.EOF
        Character = SS.PeekCharacter
        If (Character = "[") Then
            SS.EatCharacter "["
            
            Dim Number As String
            Number = ParseNumber(SS)
            Output.Add CInt(Number)
            
            SS.EatCharacter "]"
        ElseIf (Character <> "/") Then
            Dim StringData As String
            StringData = ParseString(SS)
            Output.Add StringData
        Else
            SS.EatCharacter "/"
        End If
    Loop
    Set ParseXpath = Output
End Function

Private Function ParseNumber(ByRef SS As StringStream) As String
    Debug.Assert Not (SS Is Nothing)

#If DEV Then
    Const FunctionName As String = "ParseNumber"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Number As String
    Number = SS.GetStringFromRegEx("^[1-9]\d*")
    SS.EatString Number
    ParseNumber = Number
End Function

Private Function ParseString(ByRef SS As StringStream) As String
    Debug.Assert Not (SS Is Nothing)

#If DEV Then
    Const FunctionName As String = "ParseString"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Data As String
    Data = SS.GetStringFromRegEx("[^\[\/]+")
    SS.EatString Data
    ParseString = Data
End Function
