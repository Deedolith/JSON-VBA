Attribute VB_Name = "TestModule6"
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
    Set Assert = VBA.CreateObject("Rubberduck.AssertClass")
    Set Fakes = VBA.CreateObject("Rubberduck.FakesProvider")
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

'@TestMethod("Factory")
Private Sub CreateBoolean()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JValue As Object
        Set JValue = Factory.CreateBoolean(True)
    'Act:

    'Assert:
    Assert.IsTrue TypeOf JValue Is JSON.JBoolean

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Factory")
Private Sub CreateNull()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JValue As Object
        Set JValue = Factory.CreateNull
    'Act:

    'Assert:
    Assert.IsTrue TypeOf JValue Is JSON.JNull

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Factory")
Private Sub CreateNumber()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JValue As Object
        Set JValue = Factory.CreateNumber(457.25)
    'Act:

    'Assert:
    Assert.IsTrue TypeOf JValue Is JSON.JNumber

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Factory")
Private Sub CreateString()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JValue As Object
        Set JValue = Factory.CreateString("abcdef")
    'Act:

    'Assert:
    Assert.IsTrue TypeOf JValue Is JSON.JString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Factory")
Private Sub CreateBoolean_2()
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

'@TestMethod("Factory")
Private Sub CreateNull_2()
    Const ExpectedError As Long = 438        '// Propriété ou méthode non gérée par cet objet
    On Error GoTo TestFail
    
    'Arrange:
        Dim JNull As JSON.JNull
        Set JNull = Factory.CreateNull("incorrect value")
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

'@TestMethod("Factory")
Private Sub CreateNumber_2()
    Const ExpectedError As Long = 13        '// type mismatch
    On Error GoTo TestFail
    
    'Arrange:
        Dim JNumber As JSON.JNumber
        Set JNumber = Factory.CreateNumber("incorrect value")
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

