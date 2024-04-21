VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Member"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represent a JSON object member."
'@ModuleDescription "Represent a JSON object member."
'@IgnoreModule AssignmentNotUsed, VariableNotUsed
'@Exposed
'@Folder("JSON")
Option Explicit

Private mName As String
Private mValue As JValue
Private Const ModuleName As String = "Member"

Friend Sub Create(ByVal Name As String, ByRef Value As JValue)
    Const FunctionName As String = "Create"
    Dim Logger As Logger
    Set Logger = Services.CreateErrorLogger(ModuleName, FunctionName)

    mName = Name
    Set mValue = Value
End Sub

'@Description "Return the member's name."
Public Property Get Name() As String
Attribute Name.VB_Description = "Return the member's name."
    Const FunctionName As String = "Name"
    Dim Logger As Logger
    Set Logger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Name = mName
End Property

Public Property Let Name(ByVal Value As String)
    Const FunctionName As String = "Name"
    Dim Logger As Logger
    Set Logger = Services.CreateErrorLogger(ModuleName, FunctionName)

    mName = Value
End Property

'@DefaultMember
'@Description "Return the member's value."
Public Property Get Value() As JValue
Attribute Value.VB_Description = "Return the member's value."
Attribute Value.VB_UserMemId = 0
    Const FunctionName As String = "Value"
    Dim Logger As Logger
    Set Logger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Set Value = mValue
End Property

Public Property Set Value(ByRef Data As JValue)
    Const FunctionName As String = "Value"
    Dim Logger As Logger
    Set Logger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Set mValue = Data
End Property
