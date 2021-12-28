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

