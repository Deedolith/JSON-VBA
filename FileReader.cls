VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A simple text file reader, provided for conveniance."
'@ModuleDescription "A simple text file reader, provided for conveniance."
'@IgnoreModule AssignmentNotUsed, VariableNotUsed
'@Exposed
'@Folder "Providers"
Option Explicit
Implements IReader

Private Const ModuleName As String = "FileReader"

Private mPath As String

'@Description "Constructor"
Friend Sub Create(ByVal Path As String)
Attribute Create.VB_Description = "Constructor"
    Debug.Assert Path <> vbNullString   '// path must not be empty

    Const FunctionName As String = "Create"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    mPath = Path
End Sub

'@Description "Interface implementation"
Private Function IReader_Execute() As String
Attribute IReader_Execute.VB_Description = "Interface implementation"
    Const ForReading As Long = 1

    Const FunctionName As String = "IReader_Execute"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim Fso As Object       '// Scripting.FileSystemObject
    Set Fso = VBA.CreateObject("Scripting.FileSystemObject")
    
    Dim File As Object      '// Scripting.file
    Set File = Fso.GetFile(mPath)
    
    Dim IStream As Object       '// Scripting.TextStream
    Set IStream = File.OpenAsTextStream(ForReading)
    
    Dim Data As String
    Data = IStream.ReadAll
    
    IReader_Execute = Data
End Function
