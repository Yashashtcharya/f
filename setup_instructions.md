# CATIA V5 AI Assistant - Setup Instructions

## Quick Setup Guide for PyCharm

### Step 1: Python Environment Setup

1. **Open PyCharm**
2. **Create New Project:**
   - File → New Project
   - Choose "Pure Python"
   - Select Python 3.7+ interpreter
   - Name your project (e.g., "catia-ai-assistant")

3. **Install Required Packages:**
   - Open Terminal in PyCharm (View → Tool Windows → Terminal)
   - Run these commands:
   ```bash
   pip install pywin32
   pip install requests
   ```

### Step 2: Add Project Files

1. **Copy the following files to your PyCharm project directory:**
   - `catia_ai_assistant.py` (main application)
   - `requirements.txt` (dependencies)
   - `README.md` (documentation)

2. **Or create them manually in PyCharm:**
   - Right-click project → New → Python File
   - Copy the code content from each file

### Step 3: Configure CATIA Path

1. **Open `catia_ai_assistant.py` in PyCharm**
2. **Find the `connect_to_catia` method (around line 90)**
3. **Update the CATIA path to match your installation:**
   ```python
   # Default path - update this to your CATIA installation
   catia_path = r"C:\Program Files\Dassault Systemes\B28\win_b64\code\bin\CNEXT.exe"
   ```
   
   **Common CATIA paths:**
   - `C:\Program Files\Dassault Systemes\B28\win_b64\code\bin\CNEXT.exe` (V5-6R2018)
   - `C:\Program Files\Dassault Systemes\B29\win_b64\code\bin\CNEXT.exe` (V5-6R2019)
   - `C:\Program Files\Dassault Systemes\B30\win_b64\code\bin\CNEXT.exe` (V5-6R2020)

### Step 4: Run the Application

1. **Right-click on `catia_ai_assistant.py` in PyCharm**
2. **Select "Run 'catia_ai_assistant'"**
3. **Or press Ctrl+Shift+F10**

### Step 5: Test the Application

1. **The GUI should open automatically**
2. **Click "Connect to CATIA V5"** - this will:
   - Start CATIA if not running
   - Connect to existing CATIA instance
3. **Enter a test request** like: "Create a cube with 50mm sides"
4. **Click "Generate VBA Code"**
5. **Use the generated code in CATIA**

## Debugging in PyCharm

### Setting Breakpoints
- Click in the left margin next to line numbers to set breakpoints
- Useful for debugging COM interface issues

### Viewing Variables
- Use PyCharm's debugger to inspect CATIA objects
- Watch window shows real-time variable values

### Console Output
- Check PyCharm's Run window for error messages
- Print statements will appear in the console

## Troubleshooting in PyCharm

### Import Errors
If you get import errors:
```python
# Add this at the top of the file if needed
import sys
sys.path.append(r"C:\Python39\Lib\site-packages\win32com")
```

### CATIA COM Issues
Test CATIA connection separately:
```python
# Create a test file to verify CATIA connection
import win32com.client

try:
    catia = win32com.client.Dispatch("Catia.Application")
    print("CATIA connected successfully!")
    print(f"CATIA Version: {catia.SystemConfiguration.Version}")
except Exception as e:
    print(f"CATIA connection failed: {e}")
```

### GUI Issues
Test Tkinter separately:
```python
# Create a test file to verify Tkinter
import tkinter as tk

root = tk.Tk()
root.title("Test")
tk.Label(root, text="Tkinter works!").pack()
root.mainloop()
```

## Advanced PyCharm Configuration

### Code Style
- File → Settings → Editor → Code Style → Python
- Set up PEP 8 formatting
- Enable automatic code formatting

### Virtual Environment
For better package management:
1. File → Settings → Project → Python Interpreter
2. Click gear icon → Add
3. Choose "Virtualenv Environment"
4. Install packages in isolated environment

### External Tools
Add CATIA as external tool:
1. File → Settings → Tools → External Tools
2. Add new tool:
   - Name: CATIA V5
   - Program: Path to CNEXT.exe
   - Working directory: CATIA installation folder

### Run Configuration
Create custom run configuration:
1. Run → Edit Configurations
2. Add Python configuration
3. Set environment variables if needed
4. Configure startup options

## Project Structure in PyCharm

```
Your-Project/
├── catia_ai_assistant.py     # Main application (mark as source root)
├── requirements.txt          # Dependencies
├── README.md                # Documentation
├── setup_instructions.md    # This file
├── tests/                   # Unit tests (create if needed)
│   └── test_catia_connection.py
└── templates/              # Additional VBA templates (optional)
    ├── part_templates.py
    ├── assembly_templates.py
    └── drawing_templates.py
```

## Development Workflow

1. **Code Changes**: Edit `catia_ai_assistant.py`
2. **Test**: Run with Ctrl+Shift+F10
3. **Debug**: Set breakpoints and use debugger
4. **Commit**: Use PyCharm's VCS integration
5. **Deploy**: Create executable or distribute source

## Creating Executable (Optional)

To create standalone executable:

1. **Install PyInstaller:**
   ```bash
   pip install pyinstaller
   ```

2. **Create executable:**
   ```bash
   pyinstaller --onefile --windowed catia_ai_assistant.py
   ```

3. **Find executable in `dist/` folder**

## Tips for PyCharm Development

- **Use TODO comments** for tracking development tasks
- **Enable code inspection** to catch potential issues
- **Use version control** (Git) for backup and collaboration
- **Set up code templates** for common VBA patterns
- **Configure external documentation** for CATIA API references

## Getting Help

- **PyCharm Documentation**: <https://www.jetbrains.com/help/pycharm/>
- **CATIA VBA Reference**: Available in CATIA Help
- **Python COM Programming**: Search for "Python win32com tutorials"
- **Tkinter Documentation**: <https://docs.python.org/3/library/tkinter.html>

---

**Ready to Start!** 🚀

Once you've completed these steps, you'll have a fully functional CATIA V5 AI assistant running in PyCharm. The application will help you generate VBA code for CATIA without needing any API keys!
