' CATIA V6 - Advanced Tree Name Operations
' Extended functionality for working with specification tree names

Option Explicit

' Get all tree names in the current session
Sub GetAllTreeNames()
    Dim catiaApp As Application
    Dim doc As Document
    Dim i As Integer
    Dim treeNames As String
    
    On Error GoTo ErrorHandler
    
    Set catiaApp = GetObject(, "Catia.Application")
    
    If catiaApp.Documents.Count = 0 Then
        MsgBox "No documents are currently open."
        Exit Sub
    End If
    
    treeNames = "Open Documents:" & vbCrLf & vbCrLf
    
    ' Loop through all open documents
    For i = 1 To catiaApp.Documents.Count
        Set doc = catiaApp.Documents.Item(i)
        treeNames = treeNames & i & ". " & GetDocumentTreeName(doc) & vbCrLf
    Next i
    
    MsgBox treeNames
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error getting tree names: " & Err.Description
End Sub

' Helper function to get tree name from any document type
Private Function GetDocumentTreeName(doc As Document) As String
    Dim productDoc As ProductDocument
    Dim partDoc As PartDocument
    
    On Error GoTo ErrorHandler
    
    Select Case doc.GetItem("Type")
        Case "CATProduct"
            Set productDoc = doc
            GetDocumentTreeName = productDoc.Product.Name & " (Product)"
        Case "CATPart"
            Set partDoc = doc
            GetDocumentTreeName = partDoc.Part.Name & " (Part)"
        Case Else
            GetDocumentTreeName = doc.Name & " (Other)"
    End Select
    
    Exit Function
    
ErrorHandler:
    GetDocumentTreeName = doc.Name & " (Unknown Type)"
End Function

' Get tree name with full path information
Sub GetTreeNameWithPath()
    Dim catiaApp As Application
    Dim activeDoc As Document
    Dim treeName As String
    Dim filePath As String
    Dim result As String
    
    On Error GoTo ErrorHandler
    
    Set catiaApp = GetObject(, "Catia.Application")
    Set activeDoc = catiaApp.ActiveDocument
    
    ' Get tree name
    treeName = GetTreeNameAsString()
    
    ' Get file path
    filePath = activeDoc.FullName
    
    result = "Tree Name: " & treeName & vbCrLf & _
             "File Path: " & filePath & vbCrLf & _
             "File Name: " & activeDoc.Name
    
    MsgBox result
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description
End Sub

' Check if tree name matches a specific pattern
Function IsTreeNameMatching(pattern As String) As Boolean
    Dim treeName As String
    
    treeName = GetTreeNameAsString()
    
    ' Simple pattern matching (you can enhance this with regex if needed)
    If InStr(1, UCase(treeName), UCase(pattern)) > 0 Then
        IsTreeNameMatching = True
    Else
        IsTreeNameMatching = False
    End If
End Function

' Rename the tree (if permissions allow)
Sub RenameTree(newName As String)
    Dim catiaApp As Application
    Dim activeDoc As Document
    Dim productDoc As ProductDocument
    Dim partDoc As PartDocument
    
    On Error GoTo ErrorHandler
    
    Set catiaApp = GetObject(, "Catia.Application")
    Set activeDoc = catiaApp.ActiveDocument
    
    Select Case activeDoc.GetItem("Type")
        Case "CATProduct"
            Set productDoc = activeDoc
            productDoc.Product.Name = newName
            MsgBox "Product tree renamed to: " & newName
        Case "CATPart"
            Set partDoc = activeDoc
            partDoc.Part.Name = newName
            MsgBox "Part tree renamed to: " & newName
        Case Else
            MsgBox "Cannot rename this document type"
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error renaming tree: " & Err.Description
End Sub
