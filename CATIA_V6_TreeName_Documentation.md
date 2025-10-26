# CATIA V6 Tree Name Fetching - Documentation

## Overview
This collection of VBA code provides functionality to fetch and work with tree names in CATIA V6. The code is compatible with CATIA V6R2013x and later versions.

## Files Included

### 1. CATIA_V6_GetTreeName.vba
Basic functionality for retrieving tree names from the specification tree.

**Key Functions:**
- `GetTreeName()` - Main subroutine to display tree name in a message box
- `GetTreeNameFromSpecTree()` - Alternative method using specification tree directly
- `GetTreeNameAsString()` - Function that returns tree name as string
- `ExampleUsage()` - Demonstrates how to use the functions

### 2. CATIA_V6_AdvancedTreeOperations.vba
Extended functionality for advanced tree operations.

**Key Functions:**
- `GetAllTreeNames()` - Lists all open document tree names
- `GetTreeNameWithPath()` - Gets tree name along with file path information
- `IsTreeNameMatching(pattern)` - Checks if tree name matches a pattern
- `RenameTree(newName)` - Renames the current tree (if permissions allow)

## Usage Instructions

### Basic Usage
1. Open CATIA V6
2. Open a Product (.CATProduct) or Part (.CATPart) document
3. Access the VBA editor (Tools > Macro > Visual Basic Editor)
4. Import or copy the VBA code
5. Run the desired subroutine

### Example Code Usage
```vba
' Simple tree name retrieval
Sub MyMacro()
    Dim name As String
    name = GetTreeNameAsString()
    Debug.Print "Current tree name: " & name
End Sub
```

## Document Types Supported
- **CATProduct** - Product documents (assemblies)
- **CATPart** - Part documents (individual parts)
- **Other** - Basic support for other CATIA document types

## Error Handling
All functions include comprehensive error handling that will:
- Display meaningful error messages
- Prevent crashes when no documents are open
- Handle different document types gracefully

## Requirements
- CATIA V6 (V6R2013x or later recommended)
- VBA enabled in CATIA
- Active document must be open for most functions

## Common Use Cases
1. **File Naming**: Use tree name for automated file saving
2. **Quality Control**: Verify tree names match naming conventions
3. **Batch Processing**: Process multiple documents based on tree names
4. **Logging**: Record tree names for audit trails

## Troubleshooting

### "No active document found"
- Ensure a CATIA document is open before running the code
- Check that the document is fully loaded

### "Error accessing specification tree"
- The document may not have a proper specification tree
- Try using the basic `GetTreeName()` function instead

### Permission errors when renaming
- Check if the document is read-only
- Ensure you have proper permissions to modify the document

## API References

### Key CATIA V6 Objects Used:
- `Application` - Main CATIA application object
- `Document` - Base document object
- `ProductDocument` - Product-specific document
- `PartDocument` - Part-specific document
- `SpecTree` - Specification tree object
- `Product.Name` - Product tree name property
- `Part.Name` - Part tree name property

## Notes
- Tree names in CATIA V6 are typically the same as the Product or Part name
- The specification tree root node name may differ from the Product/Part name
- Always include error handling when working with COM automation
- Test code with different document types to ensure compatibility
