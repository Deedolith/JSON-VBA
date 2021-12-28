Attribute VB_Name = "Utils"
'@Folder "JSON.Utilities"
Option Explicit

Public Enum JException
    JUnexpectedKey = vbObjectError + 1
    JUnexpectedCharacter = vbObjectError + 2
    JUnexpectedToken = vbObjectError + 3
End Enum

Public Enum JType
    JSObject
    JSArray
    JSString
    JSNumber
    JSBoolean
    JSNull
End Enum

Private Const ModuleName As String = "Utils"

Public Function GetValueAs(ByRef Value As Object, ByVal DataType As JSON.JType) As Object
#If DEV Then
    Const FunctionName As String = "GetValueAs"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim JObject As Object
    Set JObject = Services.GetValueAs(Value, DataType)
    Set GetValueAs = JObject
End Function
