VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Members"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@IgnoreModule
'@ModuleDescription "Represent all members of a JSON object."
'@Exposed
'@IgnoreModule AssignmentNotUsed, VariableNotUsed
'@Folder("JSON")
Option Explicit

Private mMembers As Collection
Private Const ModuleName As String = "Members"

Private Sub Class_Initialize()
#If DEV Then
    Const FunctionName As String = "Class_Initialize"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Set mMembers = New Collection
End Sub

'@Description "Enumerator"
'@Enumerator
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
#If DEV Then
    Const FunctionName As String = "NewEnum"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Set NewEnum = mMembers.[_NewEnum]
End Property

'@Description "Return the amount of members stored in the collection."
Public Function Count() As Long
#If DEV Then
    Const FunctionName As String = "Count"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Count = mMembers.Count
End Function

'@DefaultMember
Public Function Item(ByVal Key As String) As JSON.Pair
Attribute Item.VB_UserMemId = 0
#If DEV Then
    Const FunctionName As String = "Item"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    If (Me.HasKey(Key)) Then
        Dim Pair As JSON.Pair
        Set Pair = mMembers(Services.ToBase64(Key))
        Set Item = Pair
    Else
        Err.Raise JSON.JException.JUnexpectedKey, "Members.Item", "Unexpected key: " & Key
    End If
End Function

Public Sub Add(ByRef Pair As JSON.Pair)
    Debug.Assert Not (Pair Is Nothing)

#If DEV Then
    Const FunctionName As String = "Add"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    mMembers.Add Pair, Services.ToBase64(Pair.Name)
End Sub

Public Sub Remove(ByVal Key As String)
#If DEV Then
    Const FunctionName As String = "Remove"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    mMembers.Remove Services.ToBase64(Key)
End Sub

'@Description "Check for a key's existance (case sensitive)."
Public Function HasKey(ByVal Key As String) As Boolean
#If DEV Then
    Const FunctionName As String = "HasKey"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

On Error GoTo Error
    Dim Pair As JSON.Pair
    Set Pair = mMembers(Services.ToBase64(Key))
    HasKey = True
Exit Function
Error:
    Select Case Err.Number
    Case 5, 9      '// L'indice n'appartient pas à la sélection.
        HasKey = False
    Case Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End Select
End Function
