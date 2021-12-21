Attribute VB_Name = "Utils"
'@Folder("JSON.Utilities")
Option Explicit

Public Enum JSException
    JSUnexpectedKey = vbObjectError + 1
    JSUnexpectedCharacter = vbObjectError + 2
    JSUnexpectedToken = vbObjectError + 3
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
