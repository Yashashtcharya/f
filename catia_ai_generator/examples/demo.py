"""
Example usage of the CATIA V5 AI Code Generator
This script demonstrates various ways to use the code generator
"""

import os
import sys

# Add the src directory to the path
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'src'))

from main import AICodeGenerator, CodeRequest

def demonstrate_code_generation():
    """Demonstrate various code generation scenarios"""
    
    print("🔧 CATIA V5 AI Code Generator - Examples\n")
    
    # Initialize the generator
    generator = AICodeGenerator()
    
    examples = [
        {
            "name": "Basic Part Creation",
            "description": "Create a new part document and set the part number to 'DEMO_PART'",
            "language": "VBA",
            "complexity": "basic"
        },
        {
            "name": "Sketch with Circle",
            "description": "Create a sketch on the XY plane and draw a circle with radius 25 mm",
            "language": "VBA",
            "complexity": "intermediate"
        },
        {
            "name": "Python Automation",
            "description": "Create a Python script to batch process all parts in a directory and extract their mass properties",
            "language": "Python",
            "complexity": "advanced"
        },
        {
            "name": "Assembly Creation",
            "description": "Create an assembly document and insert two existing parts with positioning constraints",
            "language": "VBA",
            "complexity": "advanced"
        }
    ]
    
    for i, example in enumerate(examples, 1):
        print(f"{'='*60}")
        print(f"EXAMPLE {i}: {example['name']}")
        print(f"{'='*60}")
        print(f"Description: {example['description']}")
        print(f"Language: {example['language']}")
        print(f"Complexity: {example['complexity']}")
        print("\nGenerated Code:")
        print("-" * 40)
        
        # Create request
        request = CodeRequest(
            description=example['description'],
            language=example['language'],
            complexity=example['complexity']
        )
        
        # Generate code (using templates since we may not have AI API)
        generated_code = generator.generate_template_code(request)
        print(generated_code)
        print("\n")

if __name__ == "__main__":
    demonstrate_code_generation()
