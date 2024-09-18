Attribute VB_Name = "Test_JsonString"
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

'@TestMethod("Non-catégorisés")
Private Sub Instanciation_Success()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream("""SomeData""")

        Dim Json As JsonString
        Set Json = Service.CreateJsonString(Stream)
    'Act:

    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Non-catégorisés")
Private Sub Instanciation_Failure()
    Const ExpectedError As Long = JsonExceptionUnexpectedToken
    On Error GoTo TestFail
    
    'Arrange:
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream("Some\kData")

        Dim Json As JsonString
        Set Json = Service.CreateJsonString(Stream)

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

'@TestMethod("Non-catégorisés")
Private Sub ToString()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream("""SomeData""")

        Dim Json As JsonString
        Set Json = Service.CreateJsonString(Stream)
        
    'Act:
        Const Expected As String = """SomeData"""
        
        Dim Observed As String
        Observed = Json.ToString

    'Assert:
    Assert.IsTrue Observed = Expected, "Expected result is """ & Expected & """ but """ & Observed & """ was found instead."

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Non-catégorisés")
Private Sub ToString_Interface()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream("""SomeData""")
        
        Dim Json As IJson
        Set Json = Service.CreateJsonString(Stream)
        
    'Act:
        Const Expected As String = """SomeData"""
        
        Dim Observed As String
        Observed = Json.ToString

    'Assert:
    Assert.IsTrue Observed = Expected, "Expected result is """ & Expected & """ but """ & Observed & """ was found instead."

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Non-catégorisés")
Private Sub DataType()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream("""SomeData""")

        Dim Json As JsonString
        Set Json = Service.CreateJsonString(Stream)
        
    'Act:
        Const Expected As Long = JsonDataTypeString
        
        Dim Observed As Long
        Observed = Json.DataType

    'Assert:
    Assert.IsTrue Observed = Expected, "Expected result is """ & Expected & """ but """ & Observed & """ was found instead."

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Non-catégorisés")
Private Sub DataType_Interface()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream("""SomeData""")

        Dim Json As JsonString
        Set Json = Service.CreateJsonString(Stream)
        
    'Act:
        Const Expected As Long = JsonDataTypeString
        
        Dim Observed As Long
        Observed = Json.DataType

    'Assert:
    Assert.IsTrue Observed = Expected, "Expected result is """ & Expected & """ but """ & Observed & """ was found instead."

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Non-catégorisés")
Private Sub Value()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream("""SomeData""")

        Dim Json As JsonString
        Set Json = Service.CreateJsonString(Stream)
        
    'Act:
        Const Expected As String = "SomeData"
        
        Dim Observed As String
        Observed = Json.Value

    'Assert:
    Assert.IsTrue Observed = Expected, "Expected result is """ & Expected & """ but """ & Observed & """ was found instead."

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

