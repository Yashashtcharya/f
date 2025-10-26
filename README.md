# CATIA V5 AI Code Generator

A PyCharm-based Python application that integrates with CATIA V5 to automatically generate VBA code using AI, without requiring API keys.

## Features

- **CATIA V5 Integration**: Automatically connects to and opens CATIA V5
- **AI-Powered Code Generation**: Uses free AI services to generate VBA code
- **User-Friendly GUI**: Simple text input interface built with Tkinter
- **VBA Integration**: Direct integration with CATIA's VBA environment
- **Multiple Output Options**: Copy to clipboard, save to file, or insert directly into CATIA

## Prerequisites

- Windows operating system (required for CATIA V5)
- CATIA V5 installed
- Python 3.7 or higher
- PyCharm IDE (recommended)

## Installation

1. **Clone or download the project files**

2. **Install required Python packages:**
   ```bash
   pip install pywin32 requests
   ```

3. **Update CATIA path** (if needed):
   - Open `catia_ai_assistant.py`
   - Modify the `catia_path` variable in the `connect_to_catia` method to match your CATIA installation path

## Usage

1. **Run the application:**
   ```bash
   python catia_ai_assistant.py
   ```

2. **Connect to CATIA:**
   - Click "Connect to CATIA V5" button
   - The application will either connect to an existing CATIA instance or start a new one

3. **Generate VBA Code:**
   - Enter a description of what you want to create in the text area
   - Examples:
     - "Create a rectangular part with dimensions 100x50x25mm"
     - "Make a circular sketch with radius 30mm"
     - "Create an assembly with two cylindrical parts"
   - Click "Generate VBA Code"

4. **Use the Generated Code:**
   - **Insert to CATIA**: Opens VBA editor in CATIA
   - **Copy to Clipboard**: Copies code for manual pasting
   - **Save to File**: Saves code as .bas or .txt file

## How It Works

### AI Code Generation
The application uses multiple approaches to generate code:

1. **HuggingFace API**: Uses free inference endpoints (no API key required)
2. **Fallback Templates**: Uses built-in templates when AI services are unavailable
3. **Smart Matching**: Analyzes user input to select appropriate code templates

### CATIA Integration
- Uses Windows COM interface to communicate with CATIA V5
- Automatically starts CATIA if not already running
- Provides direct VBA editor access

## Code Templates

The application includes built-in templates for common CATIA operations:

- **Part Creation**: Solid modeling and feature creation
- **Assembly**: Product structure and constraints
- **Drawing**: Views and dimensions
- **Sketch**: 2D geometry creation
- **Features**: Pads, pockets, holes, etc.

## Example Use Cases

### Creating a Simple Part
**Input**: "Create a cube with 50mm sides"

**Generated Code**:
```vba
Sub CreateCube()
    Dim partDocument As Document
    Set partDocument = CATIA.Documents.Add("Part")
    
    Dim part As Part
    Set part = partDocument.Part
    
    ' Create sketch for cube base
    Dim sketches As Sketches
    Set sketches = part.Bodies.Item("PartBody").Sketches
    
    Dim sketch As Sketch
    Set sketch = sketches.Add(part.OriginElements.PlaneXY)
    
    ' Add rectangle geometry (50x50mm)
    ' ... (detailed VBA code)
    
    part.Update
    MsgBox "Cube created successfully!"
End Sub
```

### Creating a Sketch
**Input**: "Make a circle with 25mm radius on XY plane"

**Generated Code**:
```vba
Sub CreateCircleSketch()
    Dim partDocument As Document
    Set partDocument = CATIA.ActiveDocument
    
    Dim part As Part
    Set part = partDocument.Part
    
    ' Create sketch on XY plane
    Dim sketch As Sketch
    Set sketch = part.Bodies.Item("PartBody").Sketches.Add(part.OriginElements.PlaneXY)
    
    ' Create circle with 25mm radius
    ' ... (detailed VBA code)
    
    sketch.CloseEdition
    part.Update
End Sub
```

## Troubleshooting

### Common Issues

1. **"CATIA V5 not found" Error**
   - Ensure CATIA V5 is installed
   - Update the `catia_path` variable with correct installation path
   - Try starting CATIA manually first

2. **COM Interface Error**
   - Install pywin32: `pip install pywin32`
   - Run as administrator if needed
   - Ensure CATIA is not running in a different user session

3. **AI Generation Fails**
   - Check internet connection
   - The app will fall back to built-in templates
   - Try rephrasing your request

4. **VBA Editor Won't Open**
   - Ensure VBA is enabled in CATIA
   - Try opening VBA editor manually: Tools → Macro → Visual Basic Editor

## Customization

### Adding New Templates
Edit the template methods in `catia_ai_assistant.py`:
- `get_part_template()`
- `get_assembly_template()`
- `get_drawing_template()`
- `get_sketch_template()`
- `get_feature_template()`

### Using Different AI Services
Modify the `generate_with_huggingface()` method to use other free AI APIs:
- OpenAI (with free tier)
- Anthropic Claude (with free tier)
- Local AI models (Ollama, GPT4All)

### GUI Customization
The GUI is built with Tkinter and can be easily modified:
- Change colors and fonts
- Add new buttons or features
- Modify layout and sizing

## File Structure

```
catia-ai-assistant/
├── catia_ai_assistant.py    # Main application file
├── requirements.txt         # Python dependencies
├── README.md               # This documentation
└── TODO_catia_ai_assistant.md  # Project progress tracker
```

## Development Notes

- The application is designed to work without API keys
- Uses free AI inference endpoints when possible
- Includes comprehensive error handling
- Provides multiple fallback options for code generation
- Fully compatible with PyCharm development environment

## Contributing

Feel free to contribute by:
- Adding new code templates
- Improving AI integration
- Enhancing the GUI
- Adding support for more CATIA features
- Improving error handling

## License

This project is open source and available under the MIT License.
