VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A simple text file writer, provided for conveniance."
'@ModuleDescription "A simple text file writer, provided for conveniance."
'@IgnoreModule AssignmentNotUsed, VariableNotUsed
'@Exposed
'@Folder "JSON.Providers"
Option Explicit
Implements JSON.IWriter

Private mPath As String
#If DEV Then
    Private Const ModuleName As String = "FileReader"
#End If

'@Description "Constructor"
Friend Sub Create(ByVal Path As String)
Attribute Create.VB_Description = "Constructor"
    Debug.Assert Path <> vbNullString   '// path must not be empty

#If DEV Then
    Const FunctionName As String = "Create"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    mPath = Path
End Sub

'@Description "Interface implementation"
Private Sub IWriter_Execute(Data As String)
Attribute IWriter_Execute.VB_Description = "Interface implementation"
    Const ForWriting As Long = 2

#If DEV Then
    Const FunctionName As String = "IWriter_Execute"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim fso As Object       '// Scripting.FileSystemObject
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    
    Dim oStream As Object       '// Scripting.TextStream
    Set oStream = fso.OpenTextFile(mPath, ForWriting, True)
    oStream.Write Data
End Sub
