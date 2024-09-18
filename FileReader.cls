VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder("JSON.DataReaders")
Option Explicit
Implements IReader

Private Const ModuleName As String = "FileReader"

Private Type TData
    Path As String
End Type

Private This As TData

Friend Sub Create(ByVal Path As String)
    Const FunctionName As String = "Create"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    This.Path = Path
End Sub

Private Function IReader_Execute() As String
    Const FunctionName As String = "Create"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Fso As Scripting.FileSystemObject
    Set Fso = CreateObject("Scripting.fileSystemObject")
    
    Dim IStream As Scripting.TextStream
    Set IStream = Fso.OpenTextFile(This.Path, ForReading)
    
    IReader_Execute = IStream.ReadAll
End Function
