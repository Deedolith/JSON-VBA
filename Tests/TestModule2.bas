Attribute VB_Name = "TestModule2"
'@IgnoreModule
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
    'cette proc?dure s'ex?cute une seule fois par module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'cette proc?dure s'ex?cute avant chaque test dans le module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'cette proc?dure s'ex?cute apr?s chaque test dans le module.
End Sub

'@TestMethod("JNull")
Private Sub Instanciation()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JNull As JSON.JNull
        Set JNull = Factory.CreateNull
    'Act:

    'Assert:
        Assert.IsTrue TypeOf JNull Is JSON.JNull

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JNull")
Private Sub DataType()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JNull As JSON.JNull
        Set JNull = Factory.CreateNull
    'Act:

    'Assert:
        Assert.AreEqual JSON.JType.JSNull, JNull.DataType

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JNull")
Private Sub ToString()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JNull As JSON.JNull
        Set JNull = Factory.CreateNull
    'Act:

    'Assert:
        Assert.AreEqual "null", JNull.ToString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JNull")
Private Sub Value()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JNull As JSON.JNull
        Set JNull = Factory.CreateNull
    'Act:

    'Assert:
        Assert.IsTrue IsNull(JNull.Value)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JNull")
Private Sub ToJSONString()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JNull As JSON.JNull
        Set JNull = Factory.CreateNull
    'Act:

    'Assert:
        Assert.AreEqual "null", JNull.ToJSONString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JNull")
Private Sub Parsing()
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("null")
    'Act:
        Dim JNull As JSON.JNull
        Set JNull = Services.CreateNull(SS)
    'Assert:
        Assert.IsTrue IsNull(JNull.Value)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JNull")
Private Sub Parsing_2()
    Const ExpectedError As Long = JSON.JSException.JSUnexpectedToken
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("incorrect value")
    'Act:
        Dim JNull As JSON.JNull
        Set JNull = Services.CreateNull(SS)
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
