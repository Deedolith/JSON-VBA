Attribute VB_Name = "TestModule7"
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

'@TestMethod("Services")
Private Sub CreateBoolean_1()
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("true")
    'Act:
        Dim JBoolean As JSON.JBoolean
        Set JBoolean = Services.CreateBoolean(SS)
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Services")
Private Sub CreateBoolean_2()
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("false")
    'Act:
        Dim JBoolean As JSON.JBoolean
        Set JBoolean = Services.CreateBoolean(SS)
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Services")
Private Sub CreateBoolean_3()
    Const ExpectedError As Long = JSON.JException.JUnexpectedToken
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("zuebclYEL")
    'Act:
        Dim JBoolean As JSON.JBoolean
        Set JBoolean = Services.CreateBoolean(SS)
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

'@TestMethod("Services")
Private Sub CreateNull()
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("null")
    'Act:
        Dim JNull As JSON.JNull
        Set JNull = Services.CreateNull(SS)
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Services")
Private Sub CreateCreateNull_3()
    Const ExpectedError As Long = JSON.JException.JUnexpectedToken
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("zuebclYEL")
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

'@TestMethod("Services")
Private Sub CreateNumber()
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("14.5")
    'Act:
        Dim JNumber As JSON.JNumber
        Set JNumber = Services.CreateNumber(SS)
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Services")
Private Sub CreateCreateNumber_2()
    Const ExpectedError As Long = JSON.JException.JUnexpectedToken
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

'@TestMethod("Services")
Private Sub CreateString()
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("""string value""")
    'Act:
        Dim JString As JSON.JString
        Set JString = Services.CreateString(SS)
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Services")
Private Sub CreateCreateString_2()
    Const ExpectedError As Long = JSON.JException.JUnexpectedCharacter
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("incorrect value")
    'Act:
        Dim JString As JSON.JString
        Set JString = Services.CreateString(SS)
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
