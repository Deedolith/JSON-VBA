VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Base class for JSON types (array, object, boolean, null, number, string)."
'@ModuleDescription "Base class for JSON types (array, object, boolean, null, number, string)."
'@Interface
'@IgnoreModule AssignmentNotUsed, VariableNotUsed
'@Exposed
'@Folder("JSON")
Option Explicit

Private mDataType As JType
Private Const ModuleName As String = "JValue"

Friend Sub Create(ByVal DataType As JType)
    Const FunctionName As String = "Create"
    Dim Logger As Logger
    Set Logger = Services.CreateErrorLogger(ModuleName, FunctionName)

    mDataType = DataType
End Sub

'@Description "Abstract class's data type."
Public Function DataType() As JType
Attribute DataType.VB_Description = "Abstract class's data type."
    Const FunctionName As String = "DataType"
    Dim Logger As Logger
    Set Logger = Services.CreateErrorLogger(ModuleName, FunctionName)

    DataType = mDataType
End Function

'@Description "Return a string representation of the abstract class."
Public Function ToString(Optional ByVal IndentMultiplier As Long = 0) As String
Attribute ToString.VB_Description = "Return a string representation of the abstract class."
End Function

