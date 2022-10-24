Attribute VB_Name = "Test_JDocument"
'@IgnoreModule AssignmentNotUsed, VariableNotUsed, UseMeaningfulName
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

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
        If (TypeOf Data Is JSON.JString) Then
            Assert.IsTrue
        Else
            Assert.IsFalse
        End If

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JDocupment")
Private Sub Query_2()
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
        Assert.IsTrue TypeOf JDocument.Query("/8/alpha") Is JSON.JString
        Assert.AreEqual "abcdefghijklmnopqrstuvwyz", Data.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
