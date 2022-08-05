Attribute VB_Name = "Test_JObject"
'@IgnoreModule SetAssignmentWithIncompatibleObjectType, ImplicitUnboundDefaultMemberAccess, IndexedDefaultMemberAccess, UseMeaningfulName
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object        '// Rubberduck.AssertClass
'@Ignore VariableNotUsed
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
'@Ignore EmptyMethod
Private Sub TestInitialize()
    'cette procédure s'exécute avant chaque test dans le module..
End Sub

'@TestCleanup
'@Ignore EmptyMethod
Private Sub TestCleanup()
    'cette procédure s'exécute après chaque test dans le module.
End Sub

'@TestMethod("JObject")
Private Sub Instanciation()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JObject As JSON.JObject
        Set JObject = Factory.CreateObject
    'Act:
        
    'Assert:
        Assert.IsTrue TypeOf JObject Is JSON.JObject

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JObject")
Private Sub Count()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JObject As JSON.JObject
        Set JObject = Factory.CreateObject
    'Act:
        
    'Assert:
        Assert.AreEqual CLng(0), JObject.Members.Count

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JObject")
Private Sub ToString()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Expected As String
        Expected = Expected & "{" & vbCrLf
        Expected = Expected & vbCrLf
        Expected = Expected & "}"

        Dim JObject As JSON.JObject
        Set JObject = Factory.CreateObject
    'Act:

    'Assert:
        Assert.AreEqual Expected, JObject.ToString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JObject")
Private Sub ToString_2()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Expected As String
        Expected = Expected & "{" & vbCrLf
        Expected = Expected & vbTab & """Member"":null" & vbCrLf
        Expected = Expected & "}"

        Dim JObject As JSON.JObject
        Set JObject = Factory.CreateObject
    'Act:
        JObject.Members.Add Factory.CreatePair("Member", Factory.CreateNull)
    'Assert:
        Assert.AreEqual Expected, JObject.ToString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JObject")
Private Sub ToString_3()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Expected As String
        Expected = Expected & "{" & vbCrLf
        Expected = Expected & vbTab & """Member"":null," & vbCrLf
        Expected = Expected & vbTab & """Member2"":null" & vbCrLf
        Expected = Expected & "}"

        Dim JObject As JSON.JObject
        Set JObject = Factory.CreateObject
    'Act:
        JObject.Members.Add Factory.CreatePair("Member", Factory.CreateNull)
        JObject.Members.Add Factory.CreatePair("Member2", Factory.CreateNull)
    'Assert:
        Assert.AreEqual Expected, JObject.ToString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JObject")
Private Sub ToJSONString()
    On Error GoTo TestFail
    
    'Arrange:
        Const Expected As String = "{}"

        Dim JObject As JSON.JObject
        Set JObject = Factory.CreateObject
    'Act:

    'Assert:
        Assert.AreEqual Expected, JObject.ToJSONString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JObject")
Private Sub ToJSONString_2()
    On Error GoTo TestFail
    
    'Arrange:
        Const Expected As String = "{""Member"":null}"

        Dim JObject As JSON.JObject
        Set JObject = Factory.CreateObject
    'Act:
        JObject.Members.Add Factory.CreatePair("Member", Factory.CreateNull)
    'Assert:
        Assert.AreEqual Expected, JObject.ToJSONString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JObject")
Private Sub ToJSONString_3()
    On Error GoTo TestFail
    
    'Arrange:
        Const Expected As String = "{""Member"":null,""Member2"":null}"

        Dim JObject As JSON.JObject
        Set JObject = Factory.CreateObject
    'Act:
        JObject.Members.Add Factory.CreatePair("Member", Factory.CreateNull)
        JObject.Members.Add Factory.CreatePair("Member2", Factory.CreateNull)
    'Assert:
        Assert.AreEqual Expected, JObject.ToJSONString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JObject")
Private Sub Add()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JObject As JSON.JObject
        Set JObject = Factory.CreateObject
    'Act:
        JObject.Members.Add Factory.CreatePair("Member", Factory.CreateNull)
    'Assert:
        Assert.AreEqual CLng(1), JObject.Members.Count
        Assert.IsTrue TypeOf JObject.Members.Item("Member") Is JSON.Pair
        Assert.IsTrue TypeOf JObject.Members.Item("Member").Value Is JSON.JNull
        Assert.IsTrue JObject.Members.HasKey("Member")
        Assert.IsFalse JObject.Members.HasKey("member")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("JObject")
Private Sub Add_2()
    Const ExpectedError As Long = 13    '// type mismatch
    On Error GoTo TestFail
    
    'Arrange:
        Dim JObject As JSON.JObject
        Set JObject = Factory.CreateObject
    'Act:
        JObject.Members.Add Factory.CreatePair("Member", VBA.CreateObject("Scripting.FileSystemObject"))

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

'@TestMethod("JObject")
Private Sub Add_3()
    Const ExpectedError As Long = 457    '// Key already exist
    On Error GoTo TestFail
    
    'Arrange:
        Dim JObject As JSON.JObject
        Set JObject = Factory.CreateObject
    'Act:
        JObject.Members.Add Factory.CreatePair("Member", Factory.CreateNull)
        JObject.Members.Add Factory.CreatePair("Member", Factory.CreateNull)

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

'@TestMethod("JObject")
Private Sub Remove()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JObject As JSON.JObject
        Set JObject = Factory.CreateObject
        JObject.Members.Add Factory.CreatePair("Member", Factory.CreateNull)
    'Act:
        JObject.Members.Remove "Member"
    'Assert:
    Assert.AreEqual CLng(0), JObject.Members.Count

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JObject")
Private Sub Remove_2()
    Const ExpectedError As Long = 5     '// Incorrect argument or procedure call
    On Error GoTo TestFail
    
    'Arrange:
        Dim JObject As JSON.JObject
        Set JObject = Factory.CreateObject
        JObject.Members.Add Factory.CreatePair("Member", Factory.CreateNull)
    'Act:
        JObject.Members.Remove "Wrong key"

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

'@TestMethod("JObject")
Private Sub Value()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JObject As JSON.JObject
        Set JObject = Factory.CreateObject
    'Act:

    'Assert:
    Assert.AreEqual "[Object]", JObject.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JObject")
Private Sub Iterator()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JObject As JSON.JObject
        Set JObject = Factory.CreateObject
        JObject.Members.Add Factory.CreatePair("Member", Factory.CreateNull)
        JObject.Members.Add Factory.CreatePair("Member2", Factory.CreateNull)
        JObject.Members.Add Factory.CreatePair("Member3", Factory.CreateNull)
        
        Dim Count As Long
        Count = 0
    'Act:
        Dim Pair As JSON.Pair
        For Each Pair In JObject.Members
            '@Ignore VariableNotUsed
            Dim Value As Object
            Set Value = Pair.Value
            Count = Count + 1
        Next
    'Assert:
        Assert.AreEqual CLng(3), Count

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JObject")
Private Sub Parse()
    On Error GoTo TestFail
    
    'Arrange:
        Dim StringStream As JSON.StringStream
        Set StringStream = Services.CreateStringStream("{ }")
        
    'Act:
        '@Ignore VariableNotUsed
        Dim JObject As JSON.JObject
        '@Ignore AssignmentNotUsed
        Set JObject = Services.CreateObject(StringStream)
    
    'Assert:
        Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JObject")
Private Sub Parse_2()
    On Error GoTo TestFail
    
    'Arrange:
        Dim StringStream As JSON.StringStream
        Set StringStream = Services.CreateStringStream("{ ""Member"":null}")
        
    'Act:
        '@Ignore VariableNotUsed
        Dim JObject As JSON.JObject
        '@Ignore AssignmentNotUsed
        Set JObject = Services.CreateObject(StringStream)
    
    'Assert:
        Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JObject")
Private Sub Parse_3()
    Const ExpectedError As Long = JSON.JException.JUnexpectedCharacter
    On Error GoTo TestFail
    
    'Arrange:
        Dim StringStream As JSON.StringStream
        Set StringStream = Services.CreateStringStream("Incorrect value")
        
    'Act:
        '@Ignore VariableNotUsed
        Dim JObject As JSON.JObject
        '@Ignore AssignmentNotUsed
        Set JObject = Services.CreateObject(StringStream)

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

'@TestMethod("JObject")
Private Sub SetItem()
    On Error GoTo TestFail
    
    'Arrange:
        Dim StringStream As JSON.StringStream
        Set StringStream = Services.CreateStringStream("{ ""Member"":null}")
        
        Dim JObject As JSON.JObject
        Set JObject = Services.CreateObject(StringStream)
    'Act:
        Set JObject.Members("Member").Value = Factory.CreateNumber(67.4)
    'Assert:
        Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JObject")
Private Sub SetItem_2()
    On Error GoTo TestFail
    
    'Arrange:
        Dim StringStream As JSON.StringStream
        Set StringStream = Services.CreateStringStream("{ ""Member"":null}")
        
        Dim JObject As JSON.JObject
        Set JObject = Services.CreateObject(StringStream)
    'Act:
        Set JObject.Members("Member") = Factory.CreateNumber(67.4)
    'Assert:
        Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("JObject")
Private Sub SetItem_3()
    Const ExpectedError As Long = 13        '// Type mismatch
    On Error GoTo TestFail
    
    'Arrange:
        Dim StringStream As JSON.StringStream
        Set StringStream = Services.CreateStringStream("{ ""Member"":null}")
        
        Dim JObject As JSON.JObject
        Set JObject = Services.CreateObject(StringStream)
    'Act:
        Set JObject.Members("Member").Value = VBA.CreateObject("Scripting.FileSystemObject")
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
