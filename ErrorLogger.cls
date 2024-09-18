VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder("JSON.Logs")
Option Explicit

Private Type TData
    ModuleName As String
    FunctionName As String
    Loggers As Collection
End Type

Private This As TData

Private Sub Class_Initialize()
    Set This.Loggers = New Collection
    This.Loggers.Add New LoggerExecutionWindow
    This.Loggers.Add New LoggerFile
End Sub

Friend Sub Create(ByVal ModuleName As String, ByVal FunctionName As String)
    This.ModuleName = ModuleName
    This.FunctionName = FunctionName
End Sub

Private Sub Class_Terminate()
    If (Err.Number <> 0) Then
        Dim Logger As ILogger
        For Each Logger In This.Loggers
            Logger.Execute This.ModuleName & "." & This.FunctionName & " : " & Err.Description
        Next
    End If
End Sub
