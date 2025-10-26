' CATIA V6 - Fetch Tree Name
' This code demonstrates how to retrieve the tree name from the specification tree
' Compatible with CATIA V6 (V6R2013x and later versions)

Option Explicit

Sub GetTreeName()
    ' Declare CATIA application and document objects
    Dim catiaApp As Application
    Dim activeDoc As Document
    Dim productDoc As ProductDocument
    Dim partDoc As PartDocument
    Dim treeName As String
    
    ' Error handling
    On Error GoTo ErrorHandler
    
    ' Get the active CATIA application
    Set catiaApp = GetObject(, "Catia.Application")
    
    ' Check if there's an active document
    If catiaApp.Documents.Count = 0 Then
        MsgBox "No active document found. Please open a CATIA document first."
        Exit Sub
    End If
    
    ' Get the active document
    Set activeDoc = catiaApp.ActiveDocument
    
    ' Check document type and get tree name accordingly
    Select Case activeDoc.GetItem("Type")
        Case "CATProduct"
            ' For Product documents
            Set productDoc = activeDoc
            treeName = productDoc.Product.Name
            MsgBox "Product Tree Name: " & treeName
            
        Case "CATPart"
            ' For Part documents
            Set partDoc = activeDoc
            treeName = partDoc.Part.Name
            MsgBox "Part Tree Name: " & treeName
            
        Case Else
            ' For other document types
            treeName = activeDoc.Name
            MsgBox "Document Name: " & treeName
    End Select
    
    ' Output to immediate window for debugging
    Debug.Print "Tree Name: " & treeName
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description & vbCrLf & "Error Number: " & Err.Number
End Sub

' Alternative method using the specification tree directly
Sub GetTreeNameFromSpecTree()
    Dim catiaApp As Application
    Dim activeDoc As Document
    Dim specTree As SpecTree
    Dim rootNode As SpecNode
    Dim treeName As String
    
    On Error GoTo ErrorHandler
    
    ' Get CATIA application
    Set catiaApp = GetObject(, "Catia.Application")
    
    ' Get active document
    Set activeDoc = catiaApp.ActiveDocument
    
    ' Get the specification tree
    Set specTree = activeDoc.GetItem("SpecTree")
    
    ' Get the root node of the specification tree
    Set rootNode = specTree.RootNode
    
    ' Get the tree name from the root node
    treeName = rootNode.Name
    
    MsgBox "Specification Tree Root Name: " & treeName
    Debug.Print "Specification Tree Root Name: " & treeName
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error accessing specification tree: " & Err.Description
End Sub

' Function to return tree name as string (for use in other procedures)
Function GetTreeNameAsString() As String
    Dim catiaApp As Application
    Dim activeDoc As Document
    Dim productDoc As ProductDocument
    Dim partDoc As PartDocument
    Dim treeName As String
    
    On Error GoTo ErrorHandler
    
    Set catiaApp = GetObject(, "Catia.Application")
    
    If catiaApp.Documents.Count = 0 Then
        GetTreeNameAsString = "No active document"
        Exit Function
    End If
    
    Set activeDoc = catiaApp.ActiveDocument
    
    Select Case activeDoc.GetItem("Type")
        Case "CATProduct"
            Set productDoc = activeDoc
            treeName = productDoc.Product.Name
        Case "CATPart"
            Set partDoc = activeDoc
            treeName = partDoc.Part.Name
        Case Else
            treeName = activeDoc.Name
    End Select
    
    GetTreeNameAsString = treeName
    Exit Function
    
ErrorHandler:
    GetTreeNameAsString = "Error: " & Err.Description
End Function

' Example usage in a larger procedure
Sub ExampleUsage()
    Dim treeName As String
    
    ' Get the tree name
    treeName = GetTreeNameAsString()
    
    ' Use the tree name for further operations
    Debug.Print "Working with tree: " & treeName
    
    ' You can now use treeName variable for file naming, logging, etc.
End Sub
