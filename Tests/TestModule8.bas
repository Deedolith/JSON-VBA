Attribute VB_Name = "TestModule8"
'@IgnoreModule
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

'@ModuleInitialize
Private Sub ModuleInitialize()
    'cette procédure s'exécute une seule fois par module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
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
        Dim JArray As JSON.JArray
        Set JArray = Factory.CreateArray
    'Act:
        
    'Assert:
        Assert.IsTrue TypeOf JArray Is JSON.JArray

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
        Dim JArray As JSON.JArray
        Set JArray = Factory.CreateArray
    'Act:
        
    'Assert:
        Assert.AreEqual CLng(0), JArray.Size

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

        Dim JArray As JSON.JArray
        Set JArray = Factory.CreateArray
    'Act:
        
    'Assert:
        Assert.AreEqual Expected, JArray.ToString

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

        Dim JArray As JSON.JArray
        Set JArray = Factory.CreateArray
        JArray.PushBack Factory.CreateNull
    'Act:
        
    'Assert:
        Assert.AreEqual Expected, JArray.ToString

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

        Dim JArray As JSON.JArray
        Set JArray = Factory.CreateArray
        JArray.PushBack Factory.CreateNull
        JArray.PushBack Factory.CreateNull
    'Act:
        
    'Assert:
        Assert.AreEqual Expected, JArray.ToString

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

        Dim JArray As JSON.JArray
        Set JArray = Factory.CreateArray
    'Act:
        
    'Assert:
        Assert.AreEqual Expected, JArray.ToJSONString

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

        Dim JArray As JSON.JArray
        Set JArray = Factory.CreateArray
        JArray.PushBack Factory.CreateNull
    'Act:
        
    'Assert:
        Assert.AreEqual Expected, JArray.ToJSONString

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

        Dim JArray As JSON.JArray
        Set JArray = Factory.CreateArray
        JArray.PushBack Factory.CreateNull
        JArray.PushBack Factory.CreateNull
    'Act:
        
    'Assert:
        Assert.AreEqual Expected, JArray.ToJSONString

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
        Dim JArray As JSON.JArray
        Set JArray = Factory.CreateArray
    'Act:
        JArray.PushBack Factory.CreateNull
    'Assert:
        Assert.AreEqual CLng(1), JArray.Size

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
        Dim JArray As JSON.JArray
        Set JArray = Factory.CreateArray
    'Act:
        JArray.PushBack Factory.CreateNull
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
        Dim JArray As JSON.JArray
        Set JArray = Factory.CreateArray
    'Act:
        JArray.PushBack CreateObject("Scripting.FileSystemObject")

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
        Dim JArray As JSON.JArray
        Set JArray = Factory.CreateArray
    'Act:
        JArray.PushBack Factory.CreateNull
        JArray.Remove 0
    'Assert:
        Assert.AreEqual CLng(0), JArray.Size

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
        Dim JArray As JSON.JArray
        Set JArray = Factory.CreateArray
    'Act:
        JArray.Remove 0

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
        Dim JArray As JSON.JArray
        Set JArray = Factory.CreateArray
    'Act:
        JArray.PushBack Factory.CreateNull
        JArray.SetItem 0, Factory.CreateBoolean(True)
    'Assert:
        Assert.IsTrue TypeOf JArray.GetItemAs(0, JSON.JType.JSBoolean) Is JSON.JBoolean

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
        Dim JArray As JSON.JArray
        Set JArray = Factory.CreateArray
    'Act:
        JArray.PushBack Factory.CreateNull
        JArray.PushBack Factory.CreateString("random string")
        JArray.SetItem 0, Factory.CreateBoolean(True)
    'Assert:
        Assert.IsTrue TypeOf JArray.GetItemAs(0, JSON.JType.JSBoolean) Is JSON.JBoolean
        Assert.IsTrue TypeOf JArray.GetItemAs(1, JSON.JType.JSString) Is JSON.JString

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
        Dim JArray As JSON.JArray
        Set JArray = Factory.CreateArray
    'Act:
        JArray.PushBack Factory.CreateNull
        JArray.PushBack Factory.CreateString("random string")
        JArray.PushBack Factory.CreateNumber(55.4)
        JArray.SetItem 1, Factory.CreateBoolean(True)
    'Assert:
        Assert.IsTrue TypeOf JArray.GetItemAs(0, JSON.JType.JSNull) Is JSON.JNull
        Assert.IsTrue TypeOf JArray.GetItemAs(1, JSON.JType.JSBoolean) Is JSON.JBoolean
        Assert.IsTrue TypeOf JArray.GetItemAs(2, JSON.JType.JSNumber) Is JSON.JNumber

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
        Dim JArray As JSON.JArray
        Set JArray = Factory.CreateArray
        JArray.PushBack Factory.CreateNull
    'Act:
       JArray.SetItem 0, CreateObject("Scripting.FileSystemObject")
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
        Dim JArray As JSON.JArray
        Set JArray = Factory.CreateArray
    'Act:
        JArray.PushBack Factory.CreateNull
        JArray.PushBack Factory.CreateString("random string")
        JArray.PushBack Factory.CreateNumber(55.4)
        JArray.SetItem 2, Factory.CreateBoolean(True)
    'Assert:
        Assert.IsTrue TypeOf JArray.GetItemAs(0, JSON.JType.JSNull) Is JSON.JNull
        Assert.IsTrue TypeOf JArray.GetItemAs(1, JSON.JType.JSString) Is JSON.JString
        Assert.IsTrue TypeOf JArray.GetItemAs(2, JSON.JType.JSBoolean) Is JSON.JBoolean

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
        Dim JArray As JSON.JArray
        Set JArray = Factory.CreateArray
    'Act:
        JArray.PushBack Factory.CreateNull
        JArray.PushBack Factory.CreateString("random string")
        JArray.PushBack Factory.CreateNumber(55.4)
        JArray.SetItem 0, Factory.CreateBoolean(True)
    'Assert:
        Assert.IsTrue TypeOf JArray.GetItemAs(0, JSON.JType.JSBoolean) Is JSON.JBoolean
        Assert.IsTrue TypeOf JArray.GetItemAs(1, JSON.JType.JSString) Is JSON.JString
        Assert.IsTrue TypeOf JArray.GetItemAs(2, JSON.JType.JSNumber) Is JSON.JNumber

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
        Dim JArray As JSON.JArray
        Set JArray = Factory.CreateArray
    'Act:
        JArray.PushBack Factory.CreateNull
    'Assert:
        Assert.IsTrue TypeOf JArray.GetItemAs(0, JSON.JType.JSNull) Is JSON.JNull

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
        Dim JArray As JSON.JArray
        Set JArray = Factory.CreateArray
    'Act:
        JArray.PushBack Factory.CreateNull
        Assert.IsTrue TypeOf JArray.GetItemAs(0, JSON.JType.JSBoolean) Is JSON.JBoolean

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
        Dim JArray As JSON.JArray
        Set JArray = Factory.CreateArray
    'Act:
        JArray.PushBack Factory.CreateNull
        JArray.PushBack Factory.CreateNull
        JArray.PushBack Factory.CreateNull
        JArray.PushBack Factory.CreateNull
        
        Dim Element As Object
        For Each Element In JArray
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
