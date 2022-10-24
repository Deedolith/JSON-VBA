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
'@Folder "JSON.Services"
Option Explicit

#If DEV Then
    Private Const ModuleName As String = "Factory"
#End If

'@Description "Instanciate a new FileReader object"
Public Function CreateFileReader(ByVal Path As String) As JSON.FileReader
Attribute CreateFileReader.VB_Description = "Instanciate a new FileReader object"
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

'@Description "Instanciate a new FileWriter object"
Public Function CreateFileWriter(ByVal Path As String) As JSON.FileWriter
Attribute CreateFileWriter.VB_Description = "Instanciate a new FileWriter object"
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

'@Description "Instanciate a new JBoolean object"
Public Function CreateBoolean(ByVal Value As Boolean) As JSON.JBoolean
Attribute CreateBoolean.VB_Description = "Instanciate a new JBoolean object"
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

'@Description "Instanciate a new JNull object"
Public Function CreateNull() As JSON.JNull
Attribute CreateNull.VB_Description = "Instanciate a new JNull object"
#If DEV Then
    Const FunctionName As String = "CreateNull"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim JNull As JSON.JNull
    Set JNull = New JSON.JNull
    Set CreateNull = JNull
End Function

'@Description "Instanciate a new JNumber object"
Public Function CreateNumber(ByVal Value As Double) As JSON.JNumber
Attribute CreateNumber.VB_Description = "Instanciate a new JNumber object"
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

'@Description "Instanciate a new JString object"
Public Function CreateString(ByVal Value As String) As JSON.JString
Attribute CreateString.VB_Description = "Instanciate a new JString object"
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

'@Description "Instanciate a new Jarray object"
Public Function CreateArray() As JSON.Jarray
Attribute CreateArray.VB_Description = "Instanciate a new Jarray object"
#If DEV Then
    Const FunctionName As String = "CreateArray"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Jarray As JSON.Jarray
    Set Jarray = New JSON.Jarray
    Set CreateArray = Jarray
End Function

'@Description "Instanciate a new JObject object"
Public Function CreateObject() As JSON.JObject
Attribute CreateObject.VB_Description = "Instanciate a new JObject object"
#If DEV Then
    Const FunctionName As String = "CreateObject"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim JObject As JSON.JObject
    Set JObject = New JSON.JObject
    Set CreateObject = JObject
End Function

Public Function CreatePair(ByVal Name As String, ByVal Value As Object) As JSON.Pair
    Debug.Assert Not Value Is Nothing

#If DEV Then
    Const FunctionName As String = "CreatePair"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    If ((TypeOf Value Is Jarray) Or (TypeOf Value Is JBoolean) Or (TypeOf Value Is JNull) Or (TypeOf Value Is JNumber) Or (TypeOf Value Is JString) Or (TypeOf Value Is JObject)) Then
        Dim Pair As JSON.Pair
        Set Pair = New JSON.Pair
        Pair.Name = Name
        Set Pair.Value = Value
        Set CreatePair = Pair
    Else
        Err.Raise 13, "Factory.CreatePair"        '// Type Mismatch
    End If
End Function

Public Function CreateDocument() As JSON.JDocument
#If DEV Then
    Const FunctionName As String = "CreateDocument"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim JDocument As JSON.JDocument
    Set JDocument = New JSON.JDocument
    Set CreateDocument = JDocument
End Function

