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
'@Folder "Services"
Option Explicit

Private Const ModuleName As String = "Factory"

'@Description "Instanciate a new FileReader object"
Public Function CreateFileReader(ByVal Path As String) As FileReader
Attribute CreateFileReader.VB_Description = "Instanciate a new FileReader object"
    Debug.Assert Path <> vbNullString   '// path must not be empty

    Const FunctionName As String = "CreateFileReader"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim Reader As FileReader
    Set Reader = New FileReader

    Reader.Create Path
    Set CreateFileReader = Reader
End Function

'@Description "Instanciate a new FileWriter object"
Public Function CreateFileWriter(ByVal Path As String) As FileWriter
Attribute CreateFileWriter.VB_Description = "Instanciate a new FileWriter object"
    Debug.Assert Path <> vbNullString   '// path must not be empty

    Const FunctionName As String = "CreateFileWriter"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim Writer As FileWriter
    Set Writer = New FileWriter

    Writer.Create Path
    Set CreateFileWriter = Writer
End Function

'@Description "Instanciate a new JBoolean object"
Public Function CreateBoolean(ByVal Value As Boolean) As JBoolean
Attribute CreateBoolean.VB_Description = "Instanciate a new JBoolean object"
    Const FunctionName As String = "CreateBoolean"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim JBoolean As JBoolean
    Set JBoolean = New JBoolean
    JBoolean.Value = Value
    Set CreateBoolean = JBoolean
End Function

'@Description "Instanciate a new JNull object"
Public Function CreateNull() As JNull
Attribute CreateNull.VB_Description = "Instanciate a new JNull object"
    Const FunctionName As String = "CreateNull"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim JNull As JNull
    Set JNull = New JNull
    Set CreateNull = JNull
End Function

'@Description "Instanciate a new JNumber object"
Public Function CreateNumber(ByVal Value As Double) As JNumber
Attribute CreateNumber.VB_Description = "Instanciate a new JNumber object"
    Const FunctionName As String = "CreateNumber"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim JNumber As JNumber
    Set JNumber = New JNumber
    JNumber.Value = Value
    Set CreateNumber = JNumber
End Function

'@Description "Instanciate a new JString object"
Public Function CreateString(ByVal Value As String) As JString
Attribute CreateString.VB_Description = "Instanciate a new JString object"
    Const FunctionName As String = "CreateString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim JString As JString
    Set JString = New JString
    JString.Value = Value
    Set CreateString = JString
End Function

'@Description "Instanciate a new Jarray object"
Public Function CreateArray() As JArray
Attribute CreateArray.VB_Description = "Instanciate a new Jarray object"
    Const FunctionName As String = "CreateArray"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim JArray As JArray
    Set JArray = New JArray
    Set CreateArray = JArray
End Function

'@Description "Instanciate a new JObject object"
Public Function CreateObject() As JObject
Attribute CreateObject.VB_Description = "Instanciate a new JObject object"
    Const FunctionName As String = "CreateObject"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim JObject As JObject
    Set JObject = New JObject
    Set CreateObject = JObject
End Function

Public Function CreatePair(ByVal Name As String, ByVal Value As Object) As Pair
    Debug.Assert Not Value Is Nothing

    Const FunctionName As String = "CreatePair"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    If ((TypeOf Value Is JArray) Or (TypeOf Value Is JBoolean) Or (TypeOf Value Is JNull) Or (TypeOf Value Is JNumber) Or (TypeOf Value Is JString) Or (TypeOf Value Is JObject)) Then
        Dim Pair As Pair
        Set Pair = New Pair
        Pair.Name = Name
        Set Pair.Value = Value
        Set CreatePair = Pair
    Else
        Err.Raise 13, "Factory.CreatePair"        '// Type Mismatch
    End If
End Function

Public Function CreateDocument() As JDocument
    Const FunctionName As String = "CreateDocument"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim JDocument As JDocument
    Set JDocument = New JDocument
    Set CreateDocument = JDocument
End Function

