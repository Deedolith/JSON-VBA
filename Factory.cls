VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Factory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Singleton Factory class, responsible for instanciating FileReader, FileWriter and JDocument."
'@ModuleDescription "Singleton Factory class, responsible for instanciating FileReader, FileWriter and JDocument."
'@IgnoreModule AssignmentNotUsed, VariableNotUsed
'@Exposed
'@PredeclaredId
'@Folder("JSON.Services")
Option Explicit

Private Const ModuleName As String = "Factory"

Public Function CreateFileReader(ByVal Path As String) As JSON.FileReader
    Debug.Assert Path <> vbNullString   '// path must not be empty

#If DEV Then
    Const FunctionName As String = "CreateFileReader"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Reader As JSON.FileReader
    Set Reader = New JSON.FileReader

    Reader.Create Path
    Set CreateFileReader = Reader
End Function

Public Function CreateFileWriter(ByVal Path As String) As JSON.FileWriter
    Debug.Assert Path <> vbNullString   '// path must not be empty

#If DEV Then
    Const FunctionName As String = "CreateFileWriter"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Writer As JSON.FileWriter
    Set Writer = New JSON.FileWriter

    Writer.Create Path
    Set CreateFileWriter = Writer
End Function

Public Function CreateBoolean(ByVal Value As Boolean) As JSON.JBoolean
#If DEV Then
    Const FunctionName As String = "CreateBoolean"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim JBoolean As JSON.JBoolean
    Set JBoolean = New JSON.JBoolean
    JBoolean.Value = Value
    Set CreateBoolean = JBoolean
End Function

Public Function CreateNull() As JSON.JNull
#If DEV Then
    Const FunctionName As String = "CreateNull"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim JNull As JSON.JNull
    Set JNull = New JSON.JNull
    Set CreateNull = JNull
End Function

Public Function CreateNumber(ByVal Value As Double) As JSON.JNumber
#If DEV Then
    Const FunctionName As String = "CreateNumber"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim JNumber As JSON.JNumber
    Set JNumber = New JSON.JNumber
    JNumber.Value = Value
    Set CreateNumber = JNumber
End Function

Public Function CreateString(ByVal Value As String) As JSON.JString
#If DEV Then
    Const FunctionName As String = "CreateString"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim JString As JSON.JString
    Set JString = New JSON.JString
    JString.Value = Value
    Set CreateString = JString
End Function
