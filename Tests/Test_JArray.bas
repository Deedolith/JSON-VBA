Attribute VB_Name = "Test_JArray"
'@IgnoreModule
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object        '// Rubberduck.AssertClass
Private Fakes As Object         '// Rubberduck.FakesProvider

'@ModuleInitialize
Private Sub ModuleInitialize()
    'cette procédure s'exécute une seule fois par module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'cette procédure s'exécute une seule fois par module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'cette procédure s'exécute avant chaque test dans le module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'cette procédure s'exécute après chaque test dans le module.
End Sub

'@TestMethod("JArray")
Private Sub Instanciation()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
    'Act:
        
    'Assert:
        Assert.IsTrue TypeOf Jarray Is JSON.Jarray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JArray")
Private Sub Size()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
    'Act:
        
    'Assert:
        Assert.AreEqual CLng(0), Jarray.Size

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JArray")
Private Sub ToString()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Expected As String
        Expected = Expected & "[" & vbCrLf
        Expected = Expected & vbCrLf
        Expected = Expected & "]"

        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
    'Act:
        
    'Assert:
        Assert.AreEqual Expected, Jarray.ToString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JArray")
Private Sub ToString_2()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Expected As String
        Expected = Expected & "[" & vbCrLf
        Expected = Expected & vbTab & "null" & vbCrLf
        Expected = Expected & "]"

        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
        Jarray.PushBack Factory.CreateNull
    'Act:
        
    'Assert:
        Assert.AreEqual Expected, Jarray.ToString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JArray")
Private Sub ToString_3()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Expected As String
        Expected = Expected & "[" & vbCrLf
        Expected = Expected & vbTab & "null," & vbCrLf
        Expected = Expected & vbTab & "null" & vbCrLf
        Expected = Expected & "]"

        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
        Jarray.PushBack Factory.CreateNull
        Jarray.PushBack Factory.CreateNull
    'Act:
        
    'Assert:
        Assert.AreEqual Expected, Jarray.ToString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("JArray")
Private Sub ToJSONString()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Expected As String
        Expected = Expected & "[]"

        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
    'Act:
        
    'Assert:
        Assert.AreEqual Expected, Jarray.ToJSONString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JArray")
Private Sub ToJSONString_2()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Expected As String
        Expected = Expected & "[null]"

        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
        Jarray.PushBack Factory.CreateNull
    'Act:
        
    'Assert:
        Assert.AreEqual Expected, Jarray.ToJSONString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JArray")
Private Sub ToJSONString_3()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Expected As String
        Expected = Expected & "[null,null]"

        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
        Jarray.PushBack Factory.CreateNull
        Jarray.PushBack Factory.CreateNull
    'Act:
        
    'Assert:
        Assert.AreEqual Expected, Jarray.ToJSONString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("JArray")
Private Sub Size_2()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
    'Act:
        Jarray.PushBack Factory.CreateNull
    'Assert:
        Assert.AreEqual CLng(1), Jarray.Size

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JArray")
Private Sub PushBack()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
    'Act:
        Jarray.PushBack Factory.CreateNull
    'Assert:
        Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JArray")
Private Sub PushBack_2()
    Const ExpectedError As Long = 13        '// type mismatch
    On Error GoTo TestFail
    
    'Arrange:
        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
    'Act:
        Jarray.PushBack CreateObject("Scripting.FileSystemObject")

Assert:
    Assert.Fail "L'erreur attendue ne s'est pas produite"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("JArray")
Private Sub Remove()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
    'Act:
        Jarray.PushBack Factory.CreateNull
        Jarray.Remove 0
    'Assert:
        Assert.AreEqual CLng(0), Jarray.Size

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JArray")
Private Sub Remove_2()
    Const ExpectedError As Long = 5     '// Argument or procedure call incorrect
    On Error GoTo TestFail
    
    'Arrange:
        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
    'Act:
        Jarray.Remove 0

Assert:
    Assert.Fail "L'erreur attendue ne s'est pas produite"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("JArray")
Private Sub SetItem()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
    'Act:
        Jarray.PushBack Factory.CreateNull
        Jarray.SetItem 0, Factory.CreateBoolean(True)
    'Assert:
        Assert.IsTrue TypeOf Jarray.GetItemAs(0, JSON.JType.JSBoolean) Is JSON.JBoolean

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JArray")
Private Sub SetItem_2()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
    'Act:
        Jarray.PushBack Factory.CreateNull
        Jarray.PushBack Factory.CreateString("random string")
        Jarray.SetItem 0, Factory.CreateBoolean(True)
    'Assert:
        Assert.IsTrue TypeOf Jarray.GetItemAs(0, JSON.JType.JSBoolean) Is JSON.JBoolean
        Assert.IsTrue TypeOf Jarray.GetItemAs(1, JSON.JType.JSString) Is JSON.JString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JArray")
Private Sub SetItem_3()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
    'Act:
        Jarray.PushBack Factory.CreateNull
        Jarray.PushBack Factory.CreateString("random string")
        Jarray.PushBack Factory.CreateNumber(55.4)
        Jarray.SetItem 1, Factory.CreateBoolean(True)
    'Assert:
        Assert.IsTrue TypeOf Jarray.GetItemAs(0, JSON.JType.JSNull) Is JSON.JNull
        Assert.IsTrue TypeOf Jarray.GetItemAs(1, JSON.JType.JSBoolean) Is JSON.JBoolean
        Assert.IsTrue TypeOf Jarray.GetItemAs(2, JSON.JType.JSNumber) Is JSON.JNumber

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JArray")
Private Sub SetItem_4()
    Const ExpectedError As Long = 13        '// type mismatch
    On Error GoTo TestFail
    
    'Arrange:
        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
        Jarray.PushBack Factory.CreateNull
    'Act:
       Jarray.SetItem 0, CreateObject("Scripting.FileSystemObject")
Assert:
    Assert.Fail "L'erreur attendue ne s'est pas produite"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("JArray")
Private Sub SetItem_5()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
    'Act:
        Jarray.PushBack Factory.CreateNull
        Jarray.PushBack Factory.CreateString("random string")
        Jarray.PushBack Factory.CreateNumber(55.4)
        Jarray.SetItem 2, Factory.CreateBoolean(True)
    'Assert:
        Assert.IsTrue TypeOf Jarray.GetItemAs(0, JSON.JType.JSNull) Is JSON.JNull
        Assert.IsTrue TypeOf Jarray.GetItemAs(1, JSON.JType.JSString) Is JSON.JString
        Assert.IsTrue TypeOf Jarray.GetItemAs(2, JSON.JType.JSBoolean) Is JSON.JBoolean

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JArray")
Private Sub SetItem_6()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
    'Act:
        Jarray.PushBack Factory.CreateNull
        Jarray.PushBack Factory.CreateString("random string")
        Jarray.PushBack Factory.CreateNumber(55.4)
        Jarray.SetItem 0, Factory.CreateBoolean(True)
    'Assert:
        Assert.IsTrue TypeOf Jarray.GetItemAs(0, JSON.JType.JSBoolean) Is JSON.JBoolean
        Assert.IsTrue TypeOf Jarray.GetItemAs(1, JSON.JType.JSString) Is JSON.JString
        Assert.IsTrue TypeOf Jarray.GetItemAs(2, JSON.JType.JSNumber) Is JSON.JNumber

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JArray")
Private Sub GetItemAs()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
    'Act:
        Jarray.PushBack Factory.CreateNull
    'Assert:
        Assert.IsTrue TypeOf Jarray.GetItemAs(0, JSON.JType.JSNull) Is JSON.JNull

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("JArray")
Private Sub GetItemAs_2()
    Const ExpectedError As Long = 13        '// Type mismatch
    On Error GoTo TestFail
    
    'Arrange:
        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
    'Act:
        Jarray.PushBack Factory.CreateNull
        Assert.IsTrue TypeOf Jarray.GetItemAs(0, JSON.JType.JSBoolean) Is JSON.JBoolean

Assert:
    Assert.Fail "L'erreur attendue ne s'est pas produite"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("JArray")
Private Sub Iterate()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
    'Act:
        Jarray.PushBack Factory.CreateNull
        Jarray.PushBack Factory.CreateNull
        Jarray.PushBack Factory.CreateNull
        Jarray.PushBack Factory.CreateNull
        
        Dim Element As Object
        For Each Element In Jarray
            Assert.IsTrue Element.ToString <> vbNullString
        Next
    'Assert:
        Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JArray")
Private Sub Item()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
        Jarray.PushBack Factory.CreateNull
        Jarray.PushBack Factory.CreateNull
        Jarray.PushBack Factory.CreateNull
        Jarray.PushBack Factory.CreateNull
    'Act:
        Dim Element As Object
        Set Element = Jarray(2)
    'Assert:
        Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("JArray")
Private Sub Item_2()
    Const ExpectedError As Long = 9     '// L'indice n'appartient pas à la sélection.
    On Error GoTo TestFail
    
    'Arrange:
        Dim Jarray As JSON.Jarray
        Set Jarray = Factory.CreateArray
        Jarray.PushBack Factory.CreateNull
        Jarray.PushBack Factory.CreateNull
        Jarray.PushBack Factory.CreateNull
        Jarray.PushBack Factory.CreateNull
    'Act:
        Dim Element As Object
        Set Element = Jarray(4)
Assert:
    Assert.Fail "L'erreur attendue ne s'est pas produite"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("JArray")
Private Sub Parse()
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("[ ]")
        
        Dim Jarray As JSON.Jarray
        Set Jarray = Services.CreateArray(SS)
    'Act:
    
    'Assert:
        Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JArray")
Private Sub Parse_2()
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("[ null ]")
        
        Dim Jarray As JSON.Jarray
        Set Jarray = Services.CreateArray(SS)
    'Act:
    
    'Assert:
        Assert.IsTrue TypeOf Jarray.GetItemAs(0, JSON.JType.JSNull) Is JNull

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JArray")
Private Sub Parse_4()
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("[ null, null ]")
        
        Dim Jarray As JSON.Jarray
        Set Jarray = Services.CreateArray(SS)
    'Act:
    
    'Assert:
        Assert.IsTrue TypeOf Jarray.GetItemAs(0, JSON.JType.JSNull) Is JNull

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("JArray")
Private Sub Parse_3()
    Const ExpectedError As Long = JSON.JException.JUnexpectedCharacter
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("invalid array")
        
        Dim Jarray As JSON.Jarray
        Set Jarray = Services.CreateArray(SS)
    'Act:

Assert:
    Assert.Fail "L'erreur attendue ne s'est pas produite"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub
