VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder("JSON")
Option Explicit

Private mName As String
Private mElement As Object
#If DEV Then
    Const ModuleName As String = "Pair"
#End If

'@Description "Constructor"
Friend Sub Create(ByVal Name As String, ByVal Element As Object)
Attribute Create.VB_Description = "Constructor"
#If DEV Then
    Debug.Assert Not Element Is Nothing
    Const FunctionName As String = "Create"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    If ((TypeOf Element Is Jarray) Or (TypeOf Element Is JBoolean) Or (TypeOf Element Is JNull) Or (TypeOf Element Is JNumber) Or (TypeOf Element Is JString) Or (TypeOf Element Is JObject)) Then
         mName = Name
        Set mElement = Element
    Else
        Err.Raise 13        '// Type Mismatch
    End If
End Sub

'@Description "Retrieve the pair's name"
Public Property Get Name() As String
Attribute Name.VB_Description = "Retrieve the pair's name"
#If DEV Then
    Const FunctionName As String = "Name (Get)"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Name = mName
End Property

'@Description "Set the pair's name"
Public Property Let Name(ByVal Data As String)
Attribute Name.VB_Description = "Set the pair's name"
#If DEV Then
    Const FunctionName As String = "Name (Let)"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    mName = Data
End Property

'@Description "Retrieve the pair's value"
'@DefaultMember
Public Property Get Value() As Object
Attribute Value.VB_Description = "Retrieve the pair's value"
Attribute Value.VB_UserMemId = 0
#If DEV Then
    Const FunctionName As String = "Value (Get)"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Set Value = mElement
End Property

'@Description "set the pair's value"
Public Property Set Value(ByVal Element As Object)
Attribute Value.VB_Description = "set the pair's value"
#If DEV Then
    Const FunctionName As String = "Value (Set)"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If
    If ((TypeOf Element Is Jarray) Or (TypeOf Element Is JBoolean) Or (TypeOf Element Is JNull) Or (TypeOf Element Is JNumber) Or (TypeOf Element Is JString) Or (TypeOf Element Is JObject)) Then
        Set mElement = Element
    Else
        Err.Raise 13        '// Type mismatch
    End If
End Property
