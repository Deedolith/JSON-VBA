Attribute VB_Name = "Utility"
'@Folder "Utilities"
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

Private Const ModuleName As String = "Utility"

Public Function GetValueAs(ByVal Value As JSObject, ByVal DataType As JType) As JSObject
    Const FunctionName As String = "GetValueAs"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim JObject As JSObject
    Set JObject = Services.GetValueAs(Value, DataType)
    Set GetValueAs = JObject
End Function

Public Function Version() As VersionNumber
    Dim VersionNumber As VersionNumber
    VersionNumber.Major = 1
    VersionNumber.Minor = 0
    VersionNumber.Revision = 16
    Version = VersionNumber
End Function

