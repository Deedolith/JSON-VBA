Attribute VB_Name = "Utils"
'@Folder("JSON.Utilities")
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
