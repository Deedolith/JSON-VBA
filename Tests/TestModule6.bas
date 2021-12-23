Attribute VB_Name = "TestModule6"
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

''@TestMethod("Factory")
'Private Sub CreateArray()
'    On Error GoTo TestFail
'
'    'Arrange:
'        Dim JValue As Object
'        Set JValue = Factory.CreateArray
'    'Act:
'
'    'Assert:
'    Assert.IsTrue TypeOf JValue Is JSON.JArray
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub

''@TestMethod("Factory")
'Private Sub CreateObject()
'    On Error GoTo TestFail
'
'    'Arrange:
'        Dim JValue As Object
'        Set JValue = Factory.CreateObject
'    'Act:
'
'    'Assert:
'    Assert.IsTrue TypeOf JValue Is JSON.JObject
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub

''@TestMethod("Factory")
'Private Sub CreateMember()
'    On Error GoTo TestFail
'
'    'Arrange:
'        Dim JValue As Object
'        Set JValue = Factory.CreateMember("Name", Factory.CreateNull)
'    'Act:
'
'    'Assert:
'    Assert.IsTrue TypeOf JValue Is JSON.Member
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub

''@TestMethod("Factory")
'Private Sub CreateDocument()
'    On Error GoTo TestFail
'
'    'Arrange:
'        Dim JValue As Object
'        Set JValue = Factory.CreateDocument
'    'Act:
'
'    'Assert:
'    Assert.IsTrue TypeOf JValue Is JSON.JDocument
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub


