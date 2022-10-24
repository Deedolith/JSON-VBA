Attribute VB_Name = "Test_StringStream"
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

'@TestMethod("StringStream")
Private Sub Value()
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("abcdefjhijklmnopqrstuvwxyz")
    'Act:

    'Assert:
        Assert.AreEqual "abcdefjhijklmnopqrstuvwxyz", SS.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringStream")
Private Sub PeekCharacter()
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("abcdefjhijklmnopqrstuvwxyz")
    'Act:

    'Assert:
        Assert.AreEqual "a", SS.PeekCharacter()

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringStream")
Private Sub GetStringFromRegEx()
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("abcdefghijklmnopqrstuvwxyz")
    'Act:

    'Assert:
        Assert.AreEqual "fghi", SS.GetStringFromRegEx("(fghi)")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringStream")
Private Sub DiscardSpaces()
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream(" " & vbTab & vbCr & vbLf & vbCrLf & vbNewLine & vbFormFeed & vbVerticalTab & "abcdefghijklmnopqrstuvwxyz")
    'Act:
        SS.DiscardSpaces
    'Assert:
        Assert.AreEqual "abcdefghijklmnopqrstuvwxyz", SS.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringStream")
Private Sub EatCharacter()
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("abcdefghijklmnopqrstuvwxyz")
    'Act:
        SS.EatCharacter "a"
    'Assert:
        Assert.AreEqual "bcdefghijklmnopqrstuvwxyz", SS.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringStream")
Private Sub EatCharacter_fail()
    Const ExpectedError As Long = JSON.JException.JUnexpectedCharacter
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("abcdefghijklmnopqrstuvwxyz")
    'Act:
        SS.EatCharacter "b"
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

'@TestMethod("StringStream")
Private Sub EatString()
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("abcdefghijklmnopqrstuvwxyz")
    'Act:
        SS.EatString "abcdef"
    'Assert:
    Assert.AreEqual "ghijklmnopqrstuvwxyz", SS.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringStream")
Private Sub EatString_fail()
    Const ExpectedError As Long = JSON.JException.JUnexpectedCharacter
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("abcdefghijklmnopqrstuvwxyz")
    'Act:
        SS.EatString "xyz"
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

'@TestMethod("StringStream")
Private Sub EOF()
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream(vbNullString)
    'Act:

    'Assert:
    Assert.AreEqual True, SS.EOF

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

