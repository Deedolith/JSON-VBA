VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Log")
Option Explicit

#If DEV Then
    Private mModule As String
    Private mFunction As String
    
    Friend Sub Create(ByVal ModuleName As String, ByVal FunctionName As String)
        mModule = ModuleName
        mFunction = FunctionName
    End Sub
    
    Private Sub Class_Terminate()
        If (Err.Number <> 0) Then
            Debug.Print Now & vbTab & mModule & "." & mFunction & " : " & Err.Description
        End If
    End Sub
#End If

