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
'@ModuleDescription "Represent all members of a JSON JSObject."
'@Exposed
'@IgnoreModule AssignmentNotUsed, VariableNotUsed
'@Folder("JSON")
Option Explicit

Private Const ModuleName As String = "Members"

Private mMembers As Collection

Private Sub Class_Initialize()
    Const FunctionName As String = "Class_Initialize"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Set mMembers = New Collection
End Sub

'@Description "Enumerator"
'@Enumerator
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Const FunctionName As String = "NewEnum"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Set NewEnum = mMembers.[_NewEnum]
End Property

'@Description "Return the amount of members stored in the collection."
Public Function Count() As Long
    Const FunctionName As String = "Count"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Count = mMembers.Count
End Function

'@DefaultMember
'@Description "Provide access to a single member."
Public Function Item(ByVal Key As String) As Pair
Attribute Item.VB_Description = "Provide access to a single member."
Attribute Item.VB_UserMemId = 0
    Const FunctionName As String = "Item"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    If (Me.HasKey(Key)) Then
        Dim Pair As Pair
        Set Pair = mMembers(Services.ToBase64(Key))
        Set Item = Pair
    Else
        Err.Raise JException.JUnexpectedKey, "Members.Item", "Unexpected key: " & Key
    End If
End Function

'@Description "Add a member to the collection, in the form of a key / value pair"
Public Sub Add(ByRef Pair As Pair)
    Debug.Assert Not (Pair Is Nothing)
    
    Const FunctionName As String = "Add"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    mMembers.Add Pair, Services.ToBase64(Pair.Name)
End Sub

'@Description "Remove a member from the collection."
Public Sub Remove(ByVal Key As String)
    Const FunctionName As String = "Remove"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    mMembers.Remove Services.ToBase64(Key)
End Sub

'@Description "Check for a key's existance (case sensitive)."
Public Function HasKey(ByVal Key As String) As Boolean
    Const FunctionName As String = "HasKey"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

On Error GoTo Error
    Dim Pair As Pair
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
