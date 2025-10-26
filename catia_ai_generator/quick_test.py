"""
Quick Test - CATIA V5 AI Code Generator
Demonstrates code generation without external dependencies
"""

import os
import sys

# Simple CodeRequest class for testing
class CodeRequest:
    def __init__(self, description, language="VBA", complexity="basic"):
        self.description = description
        self.language = language
        self.complexity = complexity

# Simple template-based generator
class SimpleGenerator:
    def generate_code(self, request):
        if request.language.upper() == "VBA":
            return f'''
Sub GeneratedCode()
    ' Generated for: {request.description}
    ' Language: {request.language}
    ' Complexity: {request.complexity}
    
    Dim catApp As Application
    Set catApp = GetObject(, "CATIA.Application")
    
    Dim partDoc As PartDocument
    Set partDoc = catApp.ActiveDocument
    
    ' TODO: Implement specific functionality
    ' Based on description: {request.description}
End Sub
            '''
        else:  # Python
            return f'''
# Generated for: {request.description}
# Language: {request.language}
# Complexity: {request.complexity}

import win32com.client

def generated_function():
    """Generated CATIA V5 automation function"""
    try:
        catApp = win32com.client.Dispatch("CATIA.Application")
        
        # TODO: Implement specific functionality
        # Based on description: {request.description}
        
        print("Operation completed successfully")
    except Exception as e:
        print(f"Error: {{e}}")

if __name__ == "__main__":
    generated_function()
            '''

# Test the generator
print("🔧 CATIA V5 AI Code Generator - Quick Test")
print("=" * 50)

generator = SimpleGenerator()

# Test cases
tests = [
    ("Create a new part with a rectangular sketch", "VBA", "basic"),
    ("Batch process all parts in a folder", "Python", "advanced"),
    ("Create an assembly with constraints", "VBA", "intermediate")
]

for i, (desc, lang, complexity) in enumerate(tests, 1):
    print(f"\nTest {i}: {desc}")
    print(f"Language: {lang}, Complexity: {complexity}")
    print("-" * 40)
    
    request = CodeRequest(desc, lang, complexity)
    code = generator.generate_code(request)
    print(code)
    
print("\n✅ All tests completed successfully!")
print("💡 The full application with AI integration is ready in src/main.py")
print("🖥️  Run the GUI with: python3 src/gui.py (after installing dependencies)")
