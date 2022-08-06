Attribute VB_Name = "Utils"
'@Folder "JSON.Utilities"
Option Explicit

Public Const DOUBLEQUOTE As String = """"
Public Const SEMICOLON As String = ":"

Public Enum JException
    JUnexpectedKey = vbObjectError + 1
    JUnexpectedCharacter = vbObjectError + 2
    JUnexpectedToken = vbObjectError + 3
End Enum

Public Type VersionNumber
    Major As Long
    Minor As Long
    Revision As Long
End Type

Public Enum JType
    JSObject
    JSArray
    JSString
    JSNumber
    JSBoolean
    JSNull
End Enum

#If DEV Then
    Private Const ModuleName As String = "Utils"
#End If

Public Function GetValueAs(ByVal Value As Object, ByVal DataType As JSON.JType) As Object
#If DEV Then
    Const FunctionName As String = "GetValueAs"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim JObject As Object
    Set JObject = Services.GetValueAs(Value, DataType)
    Set GetValueAs = JObject
End Function

Public Function Version() As VersionNumber
    Dim VersionNumber As JSON.VersionNumber
    VersionNumber.Major = 1
    VersionNumber.Minor = 0
    VersionNumber.Revision = 14
    Version = VersionNumber
End Function

