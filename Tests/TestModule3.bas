Attribute VB_Name = "TestModule3"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

'@ModuleInitialize
Private Sub ModuleInitialize()
    'cette proc?dure s'ex?cute une seule fois par module.
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

'@TestMethod("JNumber")
Private Sub Instanciation()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JNumber As JSON.JNumber
        Set JNumber = Factory.CreateNumber(55.4)
    'Act:
        
    'Assert:
        Assert.IsTrue TypeOf JNumber Is JSON.JNumber

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JNumber")
Private Sub DataType()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JNumber As JSON.JNumber
        Set JNumber = Factory.CreateNumber(55.4)
    'Act:

    'Assert:
        Assert.AreEqual JSON.JType.JSNumber, JNumber.DataType

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JNumber")
Private Sub Value()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JNumber As JSON.JNumber
        Set JNumber = Factory.CreateNumber(55.4)
    'Act:

    'Assert:
        Assert.AreEqual 55.4, JNumber.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JNumber")
Private Sub Assignation()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JNumber As JSON.JNumber
        Set JNumber = Factory.CreateNumber(55.4)
    'Act:
        JNumber.Value = 23.6
    'Assert:
        Assert.AreEqual 23.6, JNumber.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JNumber")
Private Sub ToString()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JNumber As JSON.JNumber
        Set JNumber = Factory.CreateNumber(55.4)
    'Act:

    'Assert:
        Assert.AreEqual CStr(55.4), JNumber.ToString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JNumber")
Private Sub ToJSONString()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JNumber As JSON.JNumber
        Set JNumber = Factory.CreateNumber(55.4)
    'Act:

    'Assert:
        Assert.AreEqual "55.4", JNumber.ToJSONString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JNumber")
Private Sub Parsing()
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("55.4")

        Dim JNumber As JSON.JNumber
        Set JNumber = Services.CreateNumber(SS)
    'Act:
    
    'Assert:
        Assert.AreEqual 55.4, JNumber.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JNumber")
Private Sub Parsing_2()
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("55.4983E2")

        Dim JNumber As JSON.JNumber
        Set JNumber = Services.CreateNumber(SS)
    'Act:
    
    'Assert:
        Assert.AreEqual 5549.83, JNumber.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JNumber")
Private Sub Parsing_3()
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("55.4983e2")

        Dim JNumber As JSON.JNumber
        Set JNumber = Services.CreateNumber(SS)
    'Act:
    
    'Assert:
        Assert.AreEqual 5549.83, JNumber.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JNumber")
Private Sub Parsing_4()
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("55.4983e-2")

        Dim JNumber As JSON.JNumber
        Set JNumber = Services.CreateNumber(SS)
    'Act:
    
    'Assert:
        Assert.AreEqual 0.554983, JNumber.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JNumber")
Private Sub Assignation_2()
    Const ExpectedError As Long = 13        '// type mismatch
    On Error GoTo TestFail
    
    'Arrange:
        Dim JNumber As JSON.JNumber
        Set JNumber = Factory.CreateNumber(10)
    'Act:
        JNumber.Value = "incorrect value"

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

'@TestMethod("JNumber")
Private Sub Parsing_5()
    Const ExpectedError As Long = JSON.JSException.JSUnexpectedToken
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("incorrect value")
    'Act:
        Dim JNumber As JSON.JNumber
        Set JNumber = Services.CreateNumber(SS)
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

'@TestMethod("JNumber")
Private Sub Instanciation_2()
    Const ExpectedError As Long = 13        '// Type mismatch
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
