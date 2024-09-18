VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoggerExecutionWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("JSON.Logs")
Option Explicit
Implements ILogger

Private Sub ILogger_Execute(ByVal Message As String)
    Debug.Print Message
End Sub
