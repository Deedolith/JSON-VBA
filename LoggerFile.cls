VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoggerFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("JSON.Logs")
Option Explicit
Implements ILogger

Private Sub ILogger_Execute(ByVal Message As String)
    Dim Fso As Scripting.FileSystemObject
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    Dim Path As String
    Path = Fso.BuildPath(ThisWorkbook.Path, "Log.txt")
    
    Dim OStream As Scripting.TextStream
    Set OStream = Fso.OpenTextFile(Path, ForAppending, True)
    
    OStream.WriteLine Format$(Now, "dd/mm/yyyy hh:mm:ss") & " : " & Message
End Sub
