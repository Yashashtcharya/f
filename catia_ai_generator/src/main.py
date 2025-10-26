"""
CATIA V5 AI Code Generator
A tool that converts natural language descriptions into CATIA V5 automation code
"""

import os
import sys
import json
from typing import Dict, List, Optional
from dataclasses import dataclass
import click
from jinja2 import Template
from openai import OpenAI
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

@dataclass
class CodeRequest:
    """Data class for code generation requests"""
    description: str
    language: str = "VBA"  # VBA or Python
    complexity: str = "basic"  # basic, intermediate, advanced
    
class CatiaCodeTemplates:
    """Template library for CATIA V5 code patterns"""
    
    VBA_TEMPLATES = {
        "part_creation": '''
Sub CreatePart()
    Dim catApp As Application
    Set catApp = GetObject(, "CATIA.Application")
    
    Dim documents As Documents
    Set documents = catApp.Documents
    
    Dim partDoc As PartDocument
    Set partDoc = documents.Add("Part")
    
    ' {{ custom_code }}
End Sub
        ''',
        
        "sketch_creation": '''
Sub CreateSketch()
    Dim catApp As Application
    Set catApp = GetObject(, "CATIA.Application")
    
    Dim partDoc As PartDocument
    Set partDoc = catApp.ActiveDocument
    
    Dim part As Part
    Set part = partDoc.Part
    
    Dim bodies As Bodies
    Set bodies = part.Bodies
    
    Dim body As Body
    Set body = bodies.Item("PartBody")
    
    Dim sketches As Sketches
    Set sketches = body.Sketches
    
    ' {{ custom_code }}
End Sub
        ''',
        
        "extrude_operation": '''
Sub CreateExtrude()
    Dim catApp As Application
    Set catApp = GetObject(, "CATIA.Application")
    
    Dim partDoc As PartDocument
    Set partDoc = catApp.ActiveDocument
    
    Dim part As Part
    Set part = partDoc.Part
    
    Dim bodies As Bodies
    Set bodies = part.Bodies
    
    Dim body As Body
    Set body = bodies.Item("PartBody")
    
    ' {{ custom_code }}
End Sub
        '''
    }
    
    PYTHON_TEMPLATES = {
        "part_creation": '''
import win32com.client

def create_part():
    """Create a new CATIA V5 part"""
    try:
        catApp = win32com.client.Dispatch("CATIA.Application")
        documents = catApp.Documents
        part_doc = documents.Add("Part")
        
        # {{ custom_code }}
        
        return part_doc
    except Exception as e:
        print(f"Error creating part: {e}")
        return None
        ''',
        
        "sketch_creation": '''
def create_sketch(part_doc):
    """Create a sketch in CATIA V5 part"""
    try:
        part = part_doc.Part
        bodies = part.Bodies
        body = bodies.Item("PartBody")
        sketches = body.Sketches
        
        # {{ custom_code }}
        
    except Exception as e:
        print(f"Error creating sketch: {e}")
        ''',
        
        "automation_framework": '''
import win32com.client
import logging

class CATIAAutomation:
    """CATIA V5 Automation Framework"""
    
    def __init__(self):
        self.catApp = None
        self.active_doc = None
        self.connect_to_catia()
    
    def connect_to_catia(self):
        """Connect to CATIA application"""
        try:
            self.catApp = win32com.client.Dispatch("CATIA.Application")
            self.catApp.Visible = True
            logging.info("Connected to CATIA successfully")
        except Exception as e:
            logging.error(f"Failed to connect to CATIA: {e}")
    
    # {{ custom_methods }}
        '''
    }

class AICodeGenerator:
    """AI-powered code generator for CATIA V5"""
    
    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or os.getenv("OPENAI_API_KEY")
        self.client = OpenAI(api_key=self.api_key) if self.api_key else None
        self.templates = CatiaCodeTemplates()
    
    def generate_prompt(self, request: CodeRequest) -> str:
        """Generate a detailed prompt for the AI model"""
        prompt = f"""
You are an expert CATIA V5 automation developer. Generate {request.language} code for the following requirement:

Description: {request.description}
Language: {request.language}
Complexity Level: {request.complexity}

Requirements:
1. Generate clean, well-commented code
2. Include error handling where appropriate
3. Use CATIA V5 best practices
4. Make the code modular and reusable
5. Only return the code, no additional explanation

If using VBA:
- Use proper CATIA V5 object model
- Include necessary variable declarations
- Use appropriate error handling

If using Python:
- Use win32com.client for COM interface
- Include try-catch blocks
- Follow Python best practices
        """
        return prompt
    
    def generate_code_with_ai(self, request: CodeRequest) -> str:
        """Generate code using AI model"""
        if not self.client:
            return self.generate_template_code(request)
        
        try:
            prompt = self.generate_prompt(request)
            
            response = self.client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are an expert CATIA V5 automation developer."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=2000,
                temperature=0.3
            )
            
            return response.choices[0].message.content.strip()
            
        except Exception as e:
            print(f"AI generation failed: {e}")
            return self.generate_template_code(request)
    
    def generate_template_code(self, request: CodeRequest) -> str:
        """Generate code using templates (fallback method)"""
        templates = (self.templates.VBA_TEMPLATES if request.language.upper() == "VBA" 
                    else self.templates.PYTHON_TEMPLATES)
        
        # Simple keyword matching for template selection
        description_lower = request.description.lower()
        
        if "sketch" in description_lower:
            template_key = "sketch_creation"
        elif "extrude" in description_lower or "pad" in description_lower:
            template_key = "extrude_operation"
        elif "part" in description_lower:
            template_key = "part_creation"
        else:
            template_key = list(templates.keys())[0]  # Default to first template
        
        template = templates.get(template_key, templates[list(templates.keys())[0]])
        
        # Generate custom code snippet based on description
        custom_code = self.generate_custom_snippet(request, template_key)
        
        return Template(template).render(custom_code=custom_code)
    
    def generate_custom_snippet(self, request: CodeRequest, template_type: str) -> str:
        """Generate custom code snippet based on description"""
        description = request.description.lower()
        
        if template_type == "sketch_creation":
            return '''
    ' Create a new sketch
    Dim sketch As Sketch
    Set sketch = sketches.Add(bodies.Item("PartBody").OriginElements.PlaneYZ)
    
    ' Add geometric elements based on requirements
    ' TODO: Implement specific sketch geometry
            '''
        elif template_type == "part_creation":
            return '''
    ' Set part properties
    part.PartNumber = "GeneratedPart"
    part.Revision = "A"
    
    ' TODO: Add specific part creation logic
            '''
        else:
            return "' TODO: Implement specific functionality based on requirements"

@click.command()
@click.option('--description', '-d', required=True, help='Description of the code you want to generate')
@click.option('--language', '-l', default='VBA', type=click.Choice(['VBA', 'Python']), help='Programming language')
@click.option('--complexity', '-c', default='basic', type=click.Choice(['basic', 'intermediate', 'advanced']), help='Code complexity level')
@click.option('--output', '-o', help='Output file path')
@click.option('--use-ai', is_flag=True, help='Use AI model for code generation (requires OpenAI API key)')
def generate_code(description: str, language: str, complexity: str, output: Optional[str], use_ai: bool):
    """Generate CATIA V5 automation code from natural language description"""
    
    print(f"🔧 Generating {language} code for: {description}")
    print(f"📊 Complexity level: {complexity}")
    
    # Create code request
    request = CodeRequest(
        description=description,
        language=language,
        complexity=complexity
    )
    
    # Initialize code generator
    generator = AICodeGenerator()
    
    # Generate code
    if use_ai and generator.client:
        print("🤖 Using AI model for code generation...")
        generated_code = generator.generate_code_with_ai(request)
    else:
        print("📋 Using template-based code generation...")
        generated_code = generator.generate_template_code(request)
    
    # Output results
    if output:
        with open(output, 'w') as f:
            f.write(generated_code)
        print(f"💾 Code saved to: {output}")
    else:
        print("\n" + "="*50)
        print("GENERATED CODE:")
        print("="*50)
        print(generated_code)
        print("="*50)

if __name__ == "__main__":
    generate_code()
