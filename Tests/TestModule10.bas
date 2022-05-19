Attribute VB_Name = "TestModule10"
'@IgnoreModule AssignmentNotUsed, VariableNotUsed, UseMeaningfulName
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
'@Ignore EmptyMethod
Private Sub TestInitialize()
    'cette procédure s'exécute avant chaque test dans le module..
End Sub

'@TestCleanup
'@Ignore EmptyMethod
Private Sub TestCleanup()
    'cette procédure s'exécute après chaque test dans le module.
End Sub

'@TestMethod("JDocupment")
Private Sub LoadFrom()
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
Private Sub GetValueAs()
    On Error GoTo TestFail
    
    'Arrange:
        Const DataSource As String = "C:\Users\flambert\Desktop\JSON VBA\Test.json"
        Dim Reader As IReader
        Set Reader = Factory.CreateFileReader(DataSource)
        
        Dim JDocument As JSON.JDocument
        Set JDocument = Factory.CreateDocument
        JDocument.LoadFrom Reader
    'Act:
        Dim Jarray As JSON.Jarray
        Set Jarray = JDocument.GetValueAs(JSON.JType.JSArray)
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JDocupment")
Private Sub Query()
    On Error GoTo TestFail
    
    'Arrange:
        Const DataSource As String = "C:\Users\flambert\Desktop\JSON VBA\Test.json"
        Dim Reader As IReader
        Set Reader = Factory.CreateFileReader(DataSource)
        
        Dim JDocument As JSON.JDocument
        Set JDocument = Factory.CreateDocument
        JDocument.LoadFrom Reader
    'Act:
        Dim Data As Object
        Set Data = JDocument.Query("/8/alpha")
    'Assert:
        Assert.AreEqual "abcdefghijklmnopqrstuvwyz", Data.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

