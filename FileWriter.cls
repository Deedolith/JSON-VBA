VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder("JSON.DataWriters")
Option Explicit
Implements IWriter

Private Const ModuleName As String = "FileWriter"

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

Private Sub IWriter_Execute(ByVal Data As String)
    Const FunctionName As String = "IWriter_Execute"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Factory.CreateErrorLogger(ModuleName, FunctionName)
    
    Dim Fso As Scripting.FileSystemObject
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    Dim OStream As Scripting.TextStream
    Set OStream = Fso.OpenTextFile(This.Path, ForWriting, True)
    
    OStream.Write Data
End Sub
