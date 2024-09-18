VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Factory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed
'@Folder "JSON.Factories"
'@PredeclaredId
Option Explicit

Private Const ModuleName As String = "Factory"

Public Function CreateErrorLogger(ByVal ModuleName As String, ByVal FunctionName As String) As ErrorLogger
    Dim Logger As ErrorLogger
    Set Logger = New ErrorLogger
    
    Logger.Create ModuleName, FunctionName
    Set CreateErrorLogger = Logger
End Function

Public Function CreateStringReader(ByVal Data As String) As StringReader
    Const FunctionName As String = "CreateStringReader"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Me.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Reader As StringReader
    Set Reader = New StringReader
    
    Reader.Create Data
    Set CreateStringReader = Reader
End Function

Public Function CreateFileReader(ByVal Path As String) As FileReader
    Const FunctionName As String = "CreateStringReader"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Me.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Reader As FileReader
    Set Reader = New FileReader
    
    Reader.Create Path
    Set CreateFileReader = Reader
End Function

Public Function CreateFileWriter(ByVal Path As String) As FileWriter
    Const FunctionName As String = "CreateFileWriter"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Me.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Writer As FileWriter
    Set Writer = New FileWriter
    
    Writer.Create Path
    Set CreateFileWriter = Writer
End Function

Public Function CreateJsonDocument() As JsonDocument
    Const FunctionName As String = "CreateStringReader"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Me.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Document As JsonDocument
    Set Document = New JsonDocument
    
    Set CreateJsonDocument = Document
End Function
