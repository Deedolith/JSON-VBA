Attribute VB_Name = "TestModule10"
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

'@TestMethod("JDocupment")
Private Sub LoadFrom()                        'TODO Renommer le test
    On Error GoTo TestFail
    
    'Arrange:
        Const DataSource As String = "C:\Users\flambert\Desktop\JSON VBA\Test.json"
        Dim Reader As IReader
        Set Reader = Factory.CreateFileReader(DataSource)
        
        Dim JDocument As JSON.JDocument
        Set JDocument = Factory.CreateDocument
    'Act:
        JDocument.LoadFrom Reader
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JDocupment")
Private Sub GetValueAs()                        'TODO Renommer le test
    On Error GoTo TestFail
    
    'Arrange:
        Const DataSource As String = "C:\Users\flambert\Desktop\JSON VBA\Test.json"
        Dim Reader As IReader
        Set Reader = Factory.CreateFileReader(DataSource)
        
        Dim JDocument As JSON.JDocument
        Set JDocument = Factory.CreateDocument
        JDocument.LoadFrom Reader
    'Act:
        Dim JArray As JSON.JArray
        Set JArray = JDocument.GetValueAs(JSON.JType.JSArray)
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

