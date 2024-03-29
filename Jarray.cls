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
'@Folder "JSON"
'@ModuleDescription "Represent a JSON array, zero index based."
'@IgnoreModule AssignmentNotUsed, VariableNotUsed
'@Exposed
Option Explicit

Private mData As JSON.StringStream
Private mElements As Collection
#If DEV Then
    Private Const ModuleName As String = "JArray"
#End If

'@Description "Constructor"
Friend Sub Create(ByVal StringStream As JSON.StringStream)
Attribute Create.VB_Description = "Constructor"
#If DEV Then
    Debug.Assert Not (StringStream Is Nothing)      '// Data must be a valid instance
    Const FunctionName As String = "Create"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Set mData = StringStream
    Parse
End Sub

Private Sub Class_Initialize()
#If DEV Then
    Const FunctionName As String = "Class_Initialize"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Set mElements = New Collection
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
Private Sub ParseElement(ByVal StringStream As JSON.StringStream)
Attribute ParseElement.VB_Description = "Parse elements within the array"
#If DEV Then
    Debug.Assert Not (StringStream Is Nothing)
    Const FunctionName As String = "ParseElement"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    StringStream.DiscardSpaces
    mElements.Add Services.ParseValue(StringStream)
    StringStream.DiscardSpaces
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

    Size = mElements.Count
End Function

'@Description "Return an element given its index."
Public Function GetItemAs(ByVal Index As Long, ByVal DataType As JSON.JType) As Object
Attribute GetItemAs.VB_Description = "Return an element given its index."
#If DEV Then
    Debug.Assert Index >= 0
    Debug.Assert Index < mElements.Count
    Const FunctionName As String = "Item (Get)"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If
    
    Set GetItemAs = Services.GetValueAs(mElements.Item(Index + 1), DataType)
End Function

'@Description "Set an element at the given index"
Public Sub SetItem(ByVal Index As Long, ByRef Element As Object)
Attribute SetItem.VB_Description = "Set an element at the given index"
#If DEV Then
    Debug.Assert Index >= 0
    Debug.Assert Index < mElements.Count
    Const FunctionName As String = "Item (Set)"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Dim Position As Long
    Position = Index
     If ((TypeOf Element Is Jarray) Or (TypeOf Element Is JBoolean) Or (TypeOf Element Is JNull) Or (TypeOf Element Is JNumber) Or (TypeOf Element Is JString) Or (TypeOf Element Is JObject)) Then
        Position = Position + 1
        mElements.Remove Position
        If (mElements.Count = 0) Then
                '// Collection is empty, just add the element
            mElements.Add Element
        Else
            If (Index > mElements.Count) Then
                    '// add element at the end
                mElements.Add Element, After:=Position - 1
            Else
                    '// add element elsewhere
                mElements.Add Element, Before:=Position
            End If
        End If
    Else
        Err.Raise 13        '// type mismatch
    End If
End Sub

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

    Set NewEnum = mElements.[_NewEnum]
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
    For Each Element In mElements
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
Public Function ToString(Optional ByVal IndentMultiplier As Long = 0) As String
Attribute ToString.VB_Description = "Return a human readable string representation of the object"
#If DEV Then
    Debug.Assert IndentMultiplier >= 0
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
    For Each Element In mElements
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

'@Description "The array's default value"
Public Property Get Value() As String
Attribute Value.VB_Description = "The array's default value"
#If DEV Then
    Const FunctionName As String = "Value"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Value = "[Array]"
End Property

'@Description "Add an element at the back of the array"
Public Sub PushBack(ByRef Element As Object)
Attribute PushBack.VB_Description = "Add an element at the back of the array"
#If DEV Then
    Const FunctionName As String = "PushBack"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    If ((TypeOf Element Is Jarray) Or (TypeOf Element Is JBoolean) Or (TypeOf Element Is JNull) Or (TypeOf Element Is JNumber) Or (TypeOf Element Is JString) Or (TypeOf Element Is JObject)) Then
        mElements.Add Element
    Else
        Err.Raise 13        '// type mismatch
    End If
End Sub

'@Description "Remove an element as the given index"
Public Sub Remove(ByVal Index As Long)
Attribute Remove.VB_Description = "Remove an element as the given index"
#If DEV Then
    Debug.Assert Index >= 0
    Debug.Assert Index < mElements.Count
    Const FunctionName As String = "Remove"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    mElements.Remove Index + 1
End Sub

'@DefaultMember
Public Function Item(ByVal Index As Long) As Object
Attribute Item.VB_UserMemId = 0
#If DEV Then
    Debug.Assert Index >= 0
    Debug.Assert Index < mElements.Count
    Const FunctionName As String = "Item"
    Dim Logger As JSON.Logger
    Set Logger = Services.CreateLogger(Services.LibraryName & "." & ModuleName, FunctionName)
#End If

    Set Item = mElements.Item(Index + 1)
End Function
