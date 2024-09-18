Attribute VB_Name = "Test_JsonObject"
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
        Set Stream = Service.CreateStringStream("{ ""One"" : 1, ""Two"" : 2 , ""Tree"" :  3 , ""TREE"" : 4 }")
        
        Dim Json As JsonObject
        Set Json = Service.CreateJsonObject(Stream)
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
        Set Stream = Service.CreateStringStream("{ ""One"" : 1, ""Two"" : 2 , Tree :  3 , ""TREE"" : 4 }")
        
        Dim Json As JsonObject
        Set Json = Service.CreateJsonObject(Stream)
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
Private Sub Count()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream("{ ""One"" : 1, ""Two"" : 2 , ""Tree"" :  3 , ""TREE"" : 4 }")
        
        Dim Json As JsonObject
        Set Json = Service.CreateJsonObject(Stream)
    'Act:
        Const Expected As Long = 4
        
        Dim Observed As Long
        Observed = Json.Count

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
        Set Stream = Service.CreateStringStream("{ ""One"" : 1, ""Two"" : 2 , ""Tree"" :  3 , ""TREE"" : 4 }")
        
        Dim Json As JsonObject
        Set Json = Service.CreateJsonObject(Stream)
        
    'Act:
        Const Expected As Long = JsonDataTypeObject
        
        Dim Observed As JsonDataTypeEnum
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
        Set Stream = Service.CreateStringStream("{ ""One"" : 1, ""Two"" : 2 , ""Tree"" :  3 , ""TREE"" : 4 }")
        
        Dim Json As IJson
        Set Json = Service.CreateJsonObject(Stream)
        
    'Act:
        Const Expected As Long = JsonDataTypeObject
        
        Dim Observed As JsonDataTypeEnum
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
Private Sub ToString()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream("{ ""One"" : 1, ""Two"" : 2 , ""Tree"" :  3 , ""TREE"" : 4 }")
        
        Dim Json As JsonObject
        Set Json = Service.CreateJsonObject(Stream)
        
    'Act:
        Const Expected As String = "{""One"":1,""Two"":2,""Tree"":3,""TREE"":4}"
        
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
        Set Stream = Service.CreateStringStream("{ ""One"" : 1, ""Two"" : 2 , ""Tree"" :  3 , ""TREE"" : 4 }")
        
        Dim Json As IJson
        Set Json = Service.CreateJsonObject(Stream)
        
    'Act:
        Const Expected As String = "{""One"":1,""Two"":2,""Tree"":3,""TREE"":4}"
        
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
Private Sub Value()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream("{ ""One"" : 1, ""Two"" : 2 , ""Tree"" :  3 , ""TREE"" : 4 }")
        
        Dim Json As JsonObject
        Set Json = Service.CreateJsonObject(Stream)
        
    'Act:
        Const Expected As String = "[Object]"
        
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

'@TestMethod("Non-catégorisés")
Private Sub Iteration()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream("{ ""One"" : 1, ""Two"" : 2 , ""Tree"" :  3 , ""TREE"" : 4 }")
        
        Dim Json As JsonObject
        Set Json = Service.CreateJsonObject(Stream)
        
    'Act:
        Dim Pair As Pair
        For Each Pair In Json
        Next

    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Non-catégorisés")
Private Sub MemberAccess()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream("{ ""One"" : 1, ""Two"" : 2 , ""Tree"" :  3 , ""TREE"" : 4 }")
        
        Dim Json As JsonObject
        Set Json = Service.CreateJsonObject(Stream)
        
        Dim Pair As Pair
        Set Pair = Json.Items("TREE")
        
    'Act:
        Const ExpectedKey As String = "TREE"
        Const ExpectedValue As Long = 4
        
        Dim ObservedKey As String
        ObservedKey = Pair.Key
        
        Dim Number As JsonNumber
        Set Number = Pair.Value
        
        Dim ObservedValue As Long
        ObservedValue = Number.Value
    'Assert:
    Assert.IsTrue ObservedKey = ExpectedKey And ObservedValue = ExpectedValue, "Expected result are """ & ExpectedKey & ", " & ExpectedValue & """ but """ & ObservedKey & ", " & ObservedValue & """ were found instead."

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


