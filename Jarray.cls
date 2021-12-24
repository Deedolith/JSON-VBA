VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Jarray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represent a JSON array, zero index based."
'@Folder("JSON")
'@ModuleDescription "Represent a JSON array, zero index based."
'@IgnoreModule AssignmentNotUsed, VariableNotUsed
'@Exposed
Option Explicit

Private mData As JSON.StringStream
Private mValues As Collection
Private Const ModuleName As String = "JArray"

'@Description "Constructor"
Friend Sub Create(ByRef SS As JSON.StringStream)
Attribute Create.VB_Description = "Constructor"
    Debug.Assert Not (SS Is Nothing)      '// Data must be a valid instance

#If DEV Then
    Const FunctionName As String = "Create"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Set mData = SS
    Parse
End Sub

Private Sub Class_Initialize()
#If DEV Then
    Const FunctionName As String = "Class_Initialize"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Set mValues = New Collection
End Sub

'@Description "Stream parsing"
Private Sub Parse()
Attribute Parse.VB_Description = "Stream parsing"
#If DEV Then
    Const FunctionName As String = "Parse"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    mData.EatCharacter "["
    Do
        Dim DoLoop As Boolean
        DoLoop = False

        mData.DiscardSpaces
        If (mData.PeekCharacter <> "]") Then
            ParseElement mData
            mData.DiscardSpaces
            If (mData.PeekCharacter = ",") Then
                mData.EatCharacter (",")
                DoLoop = True
            End If
        End If
    Loop While DoLoop
    mData.EatCharacter "]"
End Sub

'@Description "Parse elements within the array"
Private Sub ParseElement(ByRef SS As JSON.StringStream)
Attribute ParseElement.VB_Description = "Parse elements within the array"
'    Debug.Assert Not (SS Is Nothing)
'
'#If DEV Then
'    Const FunctionName As String = "ParseElement"
'    Dim Logger As JSON.Logger
'    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
'#End If
'
'    SS.DiscardSpaces
'    mValues.Add Services.ParseValue(SS)
'    SS.DiscardSpaces
End Sub

'@Description "Class's data type."
Public Function DataType() As JSON.JType
Attribute DataType.VB_Description = "Class's data type."
#If DEV Then
    Const FunctionName As String = "DataType"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    DataType = JSON.JType.JSArray
End Function

'@Description "Return the number of elements in the array."
Public Function Size() As Long
Attribute Size.VB_Description = "Return the number of elements in the array."
#If DEV Then
    Const FunctionName As String = "Size"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Size = mValues.Count
End Function

'@DefaultMember
'@Description "Return an element given its index."
Public Function Item(ByVal index As Long) As Object
Attribute Item.VB_Description = "Return an element given its index."
Attribute Item.VB_UserMemId = 0
    Debug.Assert index >= 0

#If DEV Then
    Const FunctionName As String = "Item"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If
On Error GoTo Error
    Set Item = mValues.Item(index + 1)
Exit Function
Error:
    Select Case Err.Number
    Case 9      '// The index does not belong to the selection
        Err.Raise Err.Number, "Jarray.Item", "The index does not belong to the selection: " & index
    Case Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpContext, Err.HelpFile
    End Select
End Function

'@Enumerator
'@Description "Enumerator"
Public Function NewEnum() As Variant
Attribute NewEnum.VB_Description = "Enumerator"
Attribute NewEnum.VB_UserMemId = -4
#If DEV Then
    Const FunctionName As String = "NewEnum"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Set NewEnum = mValues.[_NewEnum]
End Function

'@Description "Return a JSON formated string representation of the object"
Public Function ToJSONString() As String
Attribute ToJSONString.VB_Description = "Return a JSON formated string representation of the object"
#If DEV Then
    Const FunctionName As String = "ToJSONString"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Output As String
    Output = Output & "["
    
    Dim SecondLoop As Boolean
    SecondLoop = False

    Dim Element As Object
    For Each Element In mValues
        If (SecondLoop) Then
            Output = Output & ","
        End If
        Output = Output & Element.ToJSONString
        SecondLoop = True
    Next
    Output = Output & "]"
    ToJSONString = Output
End Function

'@Description "Return a human readable string representation of the object"
Private Function ToString(Optional ByVal IndentMultiplier As Long = 0) As String
Attribute ToString.VB_Description = "Return a human readable string representation of the object"
    Debug.Assert IndentMultiplier >= 0

#If DEV Then
    Const FunctionName As String = "ToString"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If
    
    Dim Tabs As String
    Dim i As Long
    For i = 0 To IndentMultiplier - 1
        Tabs = Tabs & vbTab
    Next

    Dim Output As String
    Output = Output & "[" & vbCrLf

    Dim SecondLoop As Boolean
    SecondLoop = False

    Dim Element As Object
    For Each Element In mValues
        If (SecondLoop) Then
            Output = Output & "," & vbCrLf
        End If
        Output = Output & Tabs & vbTab & Element.ToString(IndentMultiplier + 1)
        SecondLoop = True
    Next
    Output = Output & vbCrLf
    Output = Output & Tabs & "]"
    ToString = Output
End Function

'@Description "The array's default value.."
Public Property Get Value() As String
Attribute Value.VB_Description = "The array's default value.."
#If DEV Then
    Const FunctionName As String = "Value"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Value = "[Array]"
End Property
