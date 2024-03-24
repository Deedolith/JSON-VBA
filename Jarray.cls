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

Private Const ModuleName As String = "JArray"

Private mData As StringStream
Private mElements As Collection

'@Description "Constructor"
Friend Sub Create(ByVal StringStream As StringStream)
Attribute Create.VB_Description = "Constructor"
    Debug.Assert Not (StringStream Is Nothing)      '// Data must be a valid instance
    
    Const FunctionName As String = "Create"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Set mData = StringStream
    Parse
End Sub

Private Sub Class_Initialize()
    Const FunctionName As String = "Class_Initialize"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Set mElements = New Collection
End Sub

'@Description "Stream parsing"
Private Sub Parse()
Attribute Parse.VB_Description = "Stream parsing"
    Const FunctionName As String = "Parse"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

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
Private Sub ParseElement(ByVal StringStream As StringStream)
Attribute ParseElement.VB_Description = "Parse elements within the array"
    Debug.Assert Not (StringStream Is Nothing)
    
    Const FunctionName As String = "ParseElement"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    StringStream.DiscardSpaces
    mElements.Add Services.ParseValue(StringStream)
    StringStream.DiscardSpaces
End Sub

'@Description "Class's data type."
Public Function DataType() As JType
Attribute DataType.VB_Description = "Class's data type."
    Const FunctionName As String = "DataType"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    DataType = JType.JSArray
End Function

'@Description "Return the number of elements in the array."
Public Function Size() As Long
Attribute Size.VB_Description = "Return the number of elements in the array."
    Const FunctionName As String = "Size"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Size = mElements.Count
End Function

'@Description "Return an element given its index."
Public Function GetItemAs(ByVal Index As Long, ByVal DataType As JType) As Object
Attribute GetItemAs.VB_Description = "Return an element given its index."
    Debug.Assert Index >= 0
    Debug.Assert Index < mElements.Count
    
    Const FunctionName As String = "Item (Get)"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)
    
    Set GetItemAs = Services.GetValueAs(mElements.Item(Index + 1), DataType)
End Function

'@Description "Set an element at the given index"
Public Sub SetItem(ByVal Index As Long, ByRef Element As Object)
Attribute SetItem.VB_Description = "Set an element at the given index"
    Debug.Assert Index >= 0
    Debug.Assert Index < mElements.Count
    
    Const FunctionName As String = "Item (Set)"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Dim Position As Long
    Position = Index
     If ((TypeOf Element Is JArray) Or (TypeOf Element Is JBoolean) Or (TypeOf Element Is JNull) Or (TypeOf Element Is JNumber) Or (TypeOf Element Is JString) Or (TypeOf Element Is JObject)) Then
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
    Const FunctionName As String = "NewEnum"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Set NewEnum = mElements.[_NewEnum]
End Function

'@Description "Return a JSON formated string representation of the object"
Public Function ToJSONString() As String
Attribute ToJSONString.VB_Description = "Return a JSON formated string representation of the object"
    Const FunctionName As String = "ToJSONString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

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
    Debug.Assert IndentMultiplier >= 0

    Const FunctionName As String = "ToString"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)
    
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
    Const FunctionName As String = "Value"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Value = "[Array]"
End Property

'@Description "Add an element at the back of the array"
Public Sub PushBack(ByRef Element As Object)
Attribute PushBack.VB_Description = "Add an element at the back of the array"
    Const FunctionName As String = "PushBack"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    If ((TypeOf Element Is JArray) Or (TypeOf Element Is JBoolean) Or (TypeOf Element Is JNull) Or (TypeOf Element Is JNumber) Or (TypeOf Element Is JString) Or (TypeOf Element Is JObject)) Then
        mElements.Add Element
    Else
        Err.Raise 13        '// type mismatch
    End If
End Sub

'@Description "Remove an element as the given index"
Public Sub Remove(ByVal Index As Long)
Attribute Remove.VB_Description = "Remove an element as the given index"
    Debug.Assert Index >= 0
    Debug.Assert Index < mElements.Count
    
    Const FunctionName As String = "Remove"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    mElements.Remove Index + 1
End Sub

'@DefaultMember
Public Function Item(ByVal Index As Long) As Object
Attribute Item.VB_UserMemId = 0
    Debug.Assert Index >= 0
    Debug.Assert Index < mElements.Count
    
    Const FunctionName As String = "Item"
    Dim ErrorLogger As ErrorLogger
    Set ErrorLogger = Services.CreateErrorLogger(ModuleName, FunctionName)

    Set Item = mElements.Item(Index + 1)
End Function
