Attribute VB_Name = "Test_StringStream"
'@IgnoreModule
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
Private Sub Insanciation_AvecValeur()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream("SomeData")
        
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
Private Sub PeekCharacter_SansValeur()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream(vbNullString)
        
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
Private Sub PeekCharacter()                        'TODO Renommer le test
    On Error GoTo TestFail
    
    'Arrange:
        Const Value As String = "SomeData"
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream(Value)
        
        Dim Expected As String
        Expected = Left$(Value, 1)

    'Act:
        Dim Observed As String
        Observed = Stream.PeekCharacter

    'Assert:
    Assert.IsTrue Observed = Expected, "Expected result is """ & Expected & """ but """ & Observed & """ was found instead."

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Non-catégorisés")
Private Sub EatCharacter()
    On Error GoTo TestFail
    
    'Arrange:
        Const Value As String = "SomeData"
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream(Value)

    'Act:
        Dim Expected As String
        Expected = Right$(Value, Len(Value) - 1)
        
        Dim Character As String
        Character = Left$(Value, 1)
        Stream.EatCharacter Character
        
        Dim Observed As String
        Observed = Stream.Value

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
        Const Value As String = "SomeData"
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream(Value)

    'Act:
        Dim Expected As String
        Expected = Value
        
        Dim Observed As String
        Observed = Stream.Value

    'Assert:
    Assert.IsTrue Observed = Expected, "Expected result is """ & Expected & """ but """ & Observed & """ was found instead."

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Non-catégorisés")
Private Sub EOF_True()                        'TODO Renommer le test
    On Error GoTo TestFail
    
    'Arrange:
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream(vbNullString)

    'Act:
        Const Expected As Boolean = True
        
        Dim Observed As Boolean
        Observed = Stream.EOF

    'Assert:
    Assert.IsTrue Observed = Expected, "Expected result is """ & Expected & """ but """ & Observed & """ was found instead."

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Non-catégorisés")
Private Sub EOF_False()                        'TODO Renommer le test
    On Error GoTo TestFail
    
    'Arrange:
        Const Value As String = "SomeData"
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream(Value)

    'Act:
        Const Expected As Boolean = False
        
        Dim Observed As Boolean
        Observed = Stream.EOF

    'Assert:
    Assert.IsTrue Observed = Expected, "Expected result is """ & Expected & """ but """ & Observed & """ was found instead."

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Non-catégorisés")
Private Sub Match_True()
    On Error GoTo TestFail
    
    'Arrange:
        Const Value As String = "SomeData"
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream(Value)

    'Act:
        Const RegEx As String = "^Some[\D]+$"
        Const Expected As Boolean = True
        
        Dim Observed As Boolean
        Observed = Stream.Match(RegEx)
    'Assert:
    Assert.IsTrue Observed = Expected, "Expected result is """ & Expected & """ but """ & Observed & """ was found instead."

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Non-catégorisés")
Private Sub Match_False()
    On Error GoTo TestFail
    
    'Arrange:
        Const Value As String = "DataSome"
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream(Value)

    'Act:
        Const RegEx As String = "^Some[\D]+$"
        Const Expected As Boolean = False
        
        Dim Observed As Boolean
        Observed = Stream.Match(RegEx)
    'Assert:
    Assert.IsTrue Observed = Expected, "Expected result is """ & Expected & """ but """ & Observed & """ was found instead."

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Non-catégorisés")
Private Sub EatString()
    On Error GoTo TestFail
    
    'Arrange:
        Const Value As String = "SomeData"
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream(Value)
    
    'Act:
        Const Expected As String = "Data"
        
        Stream.EatString "Some"

        Dim Observed As String
        Observed = Stream.Value
    'Assert:
    Assert.IsTrue Observed = Expected, "Expected result is """ & Expected & """ but """ & Observed & """ was found instead."

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Non-catégorisés")
Private Sub EatCharacter_Exception()
    Const ExpectedError As Long = JsonExceptionUnexpectedCharacter
    On Error GoTo TestFail
    
    'Arrange:
        Const Value As String = "SomeData"
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream(Value)

    'Act:
        Stream.EatCharacter "A"

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
Private Sub EatString_Exception()
    Const ExpectedError As Long = JsonExceptionUnexpectedCharacter
    On Error GoTo TestFail
    
    'Arrange:
        Const Value As String = "SomeData"
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream(Value)

    'Act:
        Stream.EatString "Data"

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
Private Sub PeekString()
    On Error GoTo TestFail
    
    'Arrange:
        Const Value As String = "SomeData"
        Dim Stream As StringStream
        Set Stream = Service.CreateStringStream(Value)
    
    'Act:
         Const Expected As String = "Some"
        
        Dim Observed As String
        Observed = Stream.PeekString("^Some")
        
    'Assert:
    Assert.IsTrue Observed = Expected, "Expected result is """ & Expected & """ but """ & Observed & """ was found instead."

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

