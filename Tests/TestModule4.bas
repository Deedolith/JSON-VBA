Attribute VB_Name = "TestModule4"
'@IgnoreModule
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
Private Sub TestInitialize()
    'cette procédure s'exécute avant chaque test dans le module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'cette procédure s'exécute après chaque test dans le module.
End Sub

'@TestMethod("JString")
Private Sub Instanciation()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JString As JSON.JString
        Set JString = Factory.CreateString("random string")
    'Act:
        
    'Assert:
        Assert.IsTrue TypeOf JString Is JSON.JString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JString")
Private Sub DataType()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JString As JSON.JString
        Set JString = Factory.CreateString("random string")
    'Act:

    'Assert:
        Assert.AreEqual JSON.JType.JSString, JString.DataType

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JString")
Private Sub Value()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JString As JSON.JString
        Set JString = Factory.CreateString("random string")
    'Act:

    'Assert:
        Assert.AreEqual "random string", JString.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JString")
Private Sub Assignation()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JString As JSON.JString
        Set JString = Factory.CreateString("random string")
    'Act:
        JString.Value = "new value"
    'Assert:
        Assert.AreEqual "new value", JString.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JString")
Private Sub ToString()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JString As JSON.JString
        Set JString = Factory.CreateString("random string")
    'Act:

    'Assert:
        Assert.AreEqual """random string""", JString.ToString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JString")
Private Sub ToString_2()
    On Error GoTo TestFail
    
    'Arrange:
        Dim JString As JSON.JString
        Set JString = Factory.CreateString("escape: /\" & vbTab & "éà")
    'Act:
    
    'Assert:
        Assert.AreEqual """escape: /\" & vbTab & "éà""", JString.ToString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JString")
Private Sub ToJSONString()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Observed As String
        Observed = "/\""" & vbCr & vbLf & vbCrLf & vbNewLine & vbTab & vbVerticalTab & vbBack
        
        Dim Expected As String
        Expected = "\/\\\""\r\n\r\n\r\n\t\f\b"
        Dim i As Byte
        For i = 32 To 126
            Select Case i
            Case 34, 47, 92
                '// do nothing
            Case Else
                Observed = Observed & Chr$(i)
                Expected = Expected & Chr$(i)
            End Select
        Next
    'Act:
        Dim JString As JSON.JString
        Set JString = Factory.CreateString(Observed)
    'Assert:
        Assert.AreEqual """" & Expected & """", JString.ToJSONString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JString")
Private Sub ToJSONString_2()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Observed As String
        Dim Expected As String

        Dim i As Byte
        For i = 130 To 140
            Observed = Observed & ChrW$(i)
            Expected = Expected & "\u" & Right("0000" & Hex(i), 4)
        Next
    'Act:
        Dim JString As JSON.JString
        Set JString = Factory.CreateString(Observed)
    'Assert:
        Assert.AreEqual """" & Expected & """", JString.ToJSONString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JString")
Private Sub ToJSONString_3()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Observed As String
        Dim Expected As String

        Dim i As Integer
        For i = &H200 To &H210
            Observed = Observed & ChrW$(i)
            Expected = Expected & "\u" & Right("0000" & Hex(i), 4)
        Next
        Dim JString As JSON.JString
        Set JString = Factory.CreateString(Observed)
    'Act:

    'Assert:
        Assert.AreEqual """" & Expected & """", JString.ToJSONString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JString")
Private Sub Parsing()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Observed As String
        Observed = Services.Escape("/\""" & vbCr & vbLf & vbCrLf & vbNewLine & vbTab & vbVerticalTab & vbBack)
        
        Dim Expected As String
        Expected = "\/\\\""\r\n\r\n\r\n\t\f\b"
        Dim i As Byte
        For i = 32 To 126
            Select Case i
            Case 34, 47, 92
                '// do nothing
            Case Else
                Observed = Observed & Chr$(i)
                Expected = Expected & Chr$(i)
            End Select
        Next
    'Act:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("""" & Observed & """")

        Dim JString As JSON.JString
        Set JString = Services.CreateString(SS)
    'Assert:
        Assert.AreEqual """" & Expected & """", JString.ToJSONString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JString")
Private Sub Parsing_2()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Observed As String
        Dim i As Byte
        For i = 130 To 140
            Observed = Observed & "\u" & Right("0000" & Hex(i), 4)
        Next
        Dim Expected As String
        Expected = Observed
    'Act:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("""" & Observed & """")

        Dim JString As JSON.JString
        Set JString = Services.CreateString(SS)
    'Assert:
        Assert.AreEqual """" & Expected & """", JString.ToJSONString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JString")
Private Sub Parsing_3()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Observed As String
        Observed = vbNullString

        Dim i As Long
        For i = &H200 To &H210
            Observed = Observed & "\u" & Right("0000" & Hex(i), 4)
        Next

        Dim Expected As String
        Expected = Observed
    'Act:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("""" & Observed & """")

        Dim JString As JSON.JString
        Set JString = Services.CreateString(SS)
    'Assert:
        Assert.AreEqual """" & Expected & """", JString.ToJSONString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JString")
Private Sub Parsing_4()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Observed As String
        Observed = Services.Escape("/\""" & vbCr & vbLf & vbCrLf & vbNewLine & vbTab & vbVerticalTab & vbBack)
        
        Dim Expected As String
        Expected = "/\""" & vbCr & vbLf & vbCrLf & vbNewLine & vbTab & vbVerticalTab & vbBack

        Dim i As Byte
        For i = 32 To 126
            Select Case i
            Case 34, 47, 92
                '// do nothing
            Case Else
                Observed = Observed & Chr$(i)
                Expected = Expected & Chr$(i)
            End Select
        Next
    'Act:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("""" & Observed & """")

        Dim JString As JSON.JString
        Set JString = Services.CreateString(SS)
    'Assert:
        Assert.AreEqual """" & Expected & """", JString.ToString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JString")
Private Sub Parsing_5()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Observed As String
        Observed = vbNullString
        
        Dim Expected As String
        Expected = vbNullString

        Dim i As Byte
        For i = 130 To 140
            Observed = Observed & "\u" & Right("0000" & Hex(i), 4)
            Expected = Expected & ChrW$(i)
        Next
    'Act:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("""" & Observed & """")

        Dim JString As JSON.JString
        Set JString = Services.CreateString(SS)
    'Assert:
        Assert.AreEqual """" & Expected & """", JString.ToString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JString")
Private Sub Parsing_6()
    On Error GoTo TestFail
    
    'Arrange:
        Dim Observed As String
        Observed = vbNullString

        Dim Expected As String
        Expected = vbNullString

        Dim i As Long
        For i = &H200 To &H210
            Observed = Observed & "\u" & Right("0000" & Hex(i), 4)
            Expected = Expected & ChrW$(i)
        Next
    'Act:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("""" & Observed & """")

        Dim JString As JSON.JString
        Set JString = Services.CreateString(SS)
    'Assert:
        Assert.AreEqual """" & Expected & """", JString.ToString

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Le test a produit une erreur: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("JString")
Private Sub Parsing_7()
    Const ExpectedError As Long = JSON.JException.JUnexpectedCharacter
    On Error GoTo TestFail
    
    'Arrange:
        Dim SS As JSON.StringStream
        Set SS = Services.CreateStringStream("invalid string")

        Dim JString As JSON.JString
        Set JString = Services.CreateString(SS)
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

