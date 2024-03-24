Attribute VB_Name = "Module1"
'@Folder("JSON")
Option Explicit

Public Sub test()
    Dim Reader As IReader
    Set Reader = Factory.CreateFileReader(ThisWorkbook.Path & "\Data.json")
    
    Dim Document As JDocument
    Set Document = Factory.CreateDocument
    Document.LoadFrom Reader

    Debug.Print "---------------------------------------"
    Dim Data As JArray
    Set Data = Document.GetValueAs(JSArray)
    Debug.Print Data.ToString
End Sub

