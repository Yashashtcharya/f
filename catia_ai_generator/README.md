# CATIA V5 AI Code Generator

A powerful AI-powered tool that converts natural language descriptions into CATIA V5 automation code. Whether you're working with VBA or Python, this tool helps you quickly generate automation scripts for CATIA V5.

## Features

- **Natural Language Input**: Describe what you want to create in plain English
- **Multiple Languages**: Generate both VBA and Python code
- **AI-Powered Generation**: Uses OpenAI's models for intelligent code generation
- **Template-Based Fallback**: Works without AI using built-in templates
- **GUI Interface**: User-friendly graphical interface
- **Command Line Interface**: Perfect for automation and batch processing
- **Pre-built Templates**: Common CATIA V5 operations ready to use

## Installation

1. **Clone or download the project**
2. **Install Python dependencies**:
   ```bash
   pip install -r requirements.txt
   ```
3. **Optional: Set up OpenAI API key** (for AI-powered generation):
   - Create a `.env` file in the project root
   - Add: `OPENAI_API_KEY=your-api-key-here`

## Usage

### GUI Interface (Recommended)

Run the graphical interface:
```bash
python src/gui.py
```

Features:
- Easy-to-use interface
- Quick example buttons
- Code preview and editing
- Save generated code to files

### Command Line Interface

Generate code from command line:
```bash
python src/main.py -d "Create a sketch with a rectangle" -l VBA -c basic
```

Options:
- `-d, --description`: Description of what you want to create (required)
- `-l, --language`: Programming language (VBA or Python, default: VBA)
- `-c, --complexity`: Complexity level (basic/intermediate/advanced, default: basic)
- `-o, --output`: Save to file
- `--use-ai`: Use AI generation (requires API key)

## Examples

### Example 1: Basic Part Creation
**Input**: "Create a new part document with basic properties"

**Generated VBA Code**:
```vba
Sub CreatePart()
    Dim catApp As Application
    Set catApp = GetObject(, "CATIA.Application")
    
    Dim documents As Documents
    Set documents = catApp.Documents
    
    Dim partDoc As PartDocument
    Set partDoc = documents.Add("Part")
    
    ' Set part properties
    part.PartNumber = "GeneratedPart"
    part.Revision = "A"
End Sub
```

### Example 2: Sketch Creation
**Input**: "Create a sketch with a rectangle of 50x30 mm"

**Generated Python Code**:
```python
import win32com.client

def create_rectangle_sketch():
    """Create a sketch with rectangle"""
    try:
        catApp = win32com.client.Dispatch("CATIA.Application")
        part_doc = catApp.ActiveDocument
        part = part_doc.Part
        
        # Create sketch
        bodies = part.Bodies
        body = bodies.Item("PartBody")
        sketches = body.Sketches
        
        # Add rectangle geometry
        # Implementation specific to requirements
        
    except Exception as e:
        print(f"Error creating sketch: {e}")
```

### Example 3: Complex Assembly
**Input**: "Create an assembly with two parts and add coincidence constraints"

The tool will generate appropriate assembly creation code with constraint logic.

## Project Structure

```
catia_ai_generator/
├── src/
│   ├── main.py          # Core application and CLI
│   ├── gui.py           # GUI interface
│   └── templates.py     # Extended template library
├── templates/           # Code templates
├── examples/           # Usage examples
├── docs/              # Documentation
├── requirements.txt   # Python dependencies
└── README.md         # This file
```

## Supported Operations

### VBA Operations
- Part creation and manipulation
- Sketch creation and geometry
- Extrude and revolve operations
- Assembly creation
- Constraint application
- Parametric design
- Batch processing

### Python Operations
- COM interface automation
- Batch file processing
- Report generation
- Advanced automation frameworks
- Error handling and logging

## Configuration

### Environment Variables
Create a `.env` file for configuration:
```
OPENAI_API_KEY=your-openai-api-key
CATIA_PATH=C:\Program Files\Dassault Systemes\B27\win_b64\code\bin\CNEXT.exe
```

### Template Customization
You can extend the templates in `src/templates.py` to add your own commonly used code patterns.

## Advanced Usage

### Custom Templates
Add your own templates to the `CatiaCodeTemplates` class:
```python
CUSTOM_TEMPLATES = {
    "my_operation": '''
    ' Your custom VBA template here
    ' Use {{ variable }} for dynamic content
    '''
}
```

### Batch Processing
Use the command line interface in scripts:
```bash
#!/bin/bash
python src/main.py -d "Create part A" -l VBA -o part_a.bas
python src/main.py -d "Create part B" -l Python -o part_b.py
```

## Requirements

- Python 3.7+
- CATIA V5 (for running generated code)
- OpenAI API key (optional, for AI generation)
- Windows OS (for COM interface)

## Troubleshooting

### Common Issues

1. **"Failed to connect to CATIA"**
   - Ensure CATIA V5 is installed and running
   - Check COM registration

2. **"AI generation failed"**
   - Verify OpenAI API key is set
   - Check internet connection
   - Falls back to template generation

3. **"Import error for win32com"**
   - Install: `pip install pywin32`

### Getting Help

- Check the examples in the `examples/` directory
- Review template code in `src/templates.py`
- Use the GUI for easier debugging

## Contributing

Feel free to contribute by:
- Adding new templates
- Improving AI prompts
- Adding new features
- Reporting bugs

## License

This project is open source. Feel free to use, modify, and distribute.

## Disclaimer

This tool generates automation code for CATIA V5. Always review and test generated code before using in production environments. The authors are not responsible for any damage caused by generated code.
