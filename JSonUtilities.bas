Attribute VB_Name = "JSonUtilities"
'@Folder("JSON")
Option Explicit

Private Const ModuleName As String = "JSonUtilities"

Public Enum JsonExceptionEnum
    JsonExceptionUnexpectedKey = vbObjectError + 50
    JsonExceptionUnexpectedCharacter = vbObjectError + 51
    JsonExceptionUnexpectedToken = vbObjectError + 52
End Enum

Public Enum JsonDataTypeEnum
    JsonDataTypeObject
    jsondatatypeArray
    JsonDataTypeString
    JsonDataTypeNumber
    JsonDataTypeBoolean
    JsonDataTypeNull
End Enum

'@Description "Check if the key Key exist within the collection Col"
Public Function ExistInCollection(ByVal Key As String, ByVal Col As Object) As Boolean
Attribute ExistInCollection.VB_Description = "Check if the key Key exist within the collection Col"
    Const FunctionName As String = "ExistInCollection"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)

    ExistInCollection = ExistInCollectionByVal(Key, Col) Or ExistInCollectionByRef(Key, Col)
End Function

Private Function ExistInCollectionByVal(ByVal Key As String, ByVal Col As Object) As Boolean
On Error GoTo Error
    Const FunctionName As String = "ExistInCollectionByVal"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)

    Dim Item As Variant
    Item = Col(Key)
    ExistInCollectionByVal = True
Exit Function
Error:
    Err.Clear
    ExistInCollectionByVal = False
End Function

Private Function ExistInCollectionByRef(ByVal Key As String, ByVal Col As Object) As Boolean
On Error GoTo Error
    Const FunctionName As String = "ExistInCollectionByRef"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)

    Dim Item As Variant
    Set Item = Col(Key)
    ExistInCollectionByRef = True
Exit Function
Error:
    Err.Clear
    ExistInCollectionByRef = False
End Function

Public Function GetAs(ByVal Item As IJson, ByVal DataType As JsonDataTypeEnum) As IJson
    Const FunctionName As String = "GetAs"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Set GetAs = Service.GetAs(Item, DataType)
End Function


