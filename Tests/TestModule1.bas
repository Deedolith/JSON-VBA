Attribute VB_Name = "TestModule1"
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

'@TestMethod("JBoolean")
Private Sub Instanciation()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JBoolean As JSON.JBoolean
        Set JBoolean = Factory.CreateBoolean(True)
    'Act:
        
    'Assert:
        Assert.IsTrue TypeOf JBoolean Is JSON.JBoolean

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JBoolean")
Private Sub DataType()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JBoolean As JSON.JBoolean
        Set JBoolean = Factory.CreateBoolean(True)
    'Act:

    'Assert:
        Assert.AreEqual JSON.JType.JSBoolean, JBoolean.DataType

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JBoolean")
Private Sub Value()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JBoolean As JSON.JBoolean
        Set JBoolean = Factory.CreateBoolean(True)
    'Act:

    'Assert:
        Assert.AreEqual True, JBoolean.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JBoolean")
Private Sub Assignation()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JBoolean As JSON.JBoolean
        Set JBoolean = Factory.CreateBoolean(True)
    'Act:
        JBoolean.Value = False
    'Assert:
        Assert.AreEqual False, JBoolean.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JBoolean")
Private Sub Assignation_2()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JBoolean As JSON.JBoolean
        Set JBoolean = Factory.CreateBoolean(True)
    'Act:
        JBoolean.Value = True
    'Assert:
        Assert.AreEqual True, JBoolean.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JBoolean")
Private Sub ToString_1()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JBoolean As JSON.JBoolean
        Set JBoolean = Factory.CreateBoolean(True)
    'Act:
        
    'Assert:
        Assert.AreEqual "true", JBoolean.ToString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JBoolean")
Private Sub ToString_2()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JBoolean As JSON.JBoolean
        Set JBoolean = Factory.CreateBoolean(False)
    'Act:
        
    'Assert:
        Assert.AreEqual "false", JBoolean.ToString
        

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JBoolean")
Private Sub ToJSONString()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JBoolean As JSON.JBoolean
        Set JBoolean = Factory.CreateBoolean(True)
    'Act:

    'Assert:
        Assert.AreEqual "true", JBoolean.ToJSONString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JBoolean")
Private Sub ToJSONString_2()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JBoolean As JSON.JBoolean
        Set JBoolean = Factory.CreateBoolean(False)
    'Act:

    'Assert:
        Assert.AreEqual "false", JBoolean.ToJSONString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JBoolean")
Private Sub Parsing()
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("true")

        Dim JBoolean As JSON.JBoolean
        Set JBoolean = Services.CreateBoolean(SS)
    'Act:

    'Assert:
        Assert.AreEqual True, JBoolean.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JBoolean")
Private Sub Parsing_2()
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("false")

        Dim JBoolean As JSON.JBoolean
        Set JBoolean = Services.CreateBoolean(SS)
    'Act:

    'Assert:
        Assert.AreEqual False, JBoolean.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JBoolean")
Private Sub Assignation_3()
    Const ExpectedError As Long = 13        '// type mismatch
    On Error GoTo TestFail
    
    'Arrange:
        Dim JBoolean As JSON.JBoolean
        Set JBoolean = Factory.CreateBoolean(True)
    'Act:
        JBoolean.Value = "incorrect value"

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

'@TestMethod("JBoolean")
Private Sub Instanciation_2()
    Const ExpectedError As Long = 13        '// type mismatch
    On Error GoTo TestFail
    
    'Arrange:
        Dim JBoolean As JSON.JBoolean
        Set JBoolean = Factory.CreateBoolean("incorrect value")
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

'@TestMethod("JBoolean")
Private Sub Parsing_3()
    Const ExpectedError As Long = JSON.JException.JUnexpectedToken
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("incorrect value")
    'Act:
        Dim JBoolean As JSON.JBoolean
        Set JBoolean = Services.CreateBoolean(SS)
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
