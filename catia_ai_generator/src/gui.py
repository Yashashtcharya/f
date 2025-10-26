#!/usr/bin/env python3
"""
CATIA V5 AI Code Generator - GUI Version
A user-friendly interface for generating CATIA V5 automation code
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import os
from pathlib import Path

# Import our main generator
import sys
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from main import AICodeGenerator, CodeRequest

class CatiaCodeGeneratorGUI:
    """GUI application for CATIA V5 code generation"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("CATIA V5 AI Code Generator")
        self.root.geometry("800x700")
        
        # Initialize the code generator
        self.generator = AICodeGenerator()
        
        self.create_widgets()
        self.setup_layout()
    
    def create_widgets(self):
        """Create all GUI widgets"""
        
        # Title
        self.title_label = tk.Label(
            self.root, 
            text="CATIA V5 AI Code Generator", 
            font=("Arial", 16, "bold"),
            fg="blue"
        )
        
        # Description input
        self.desc_label = tk.Label(self.root, text="Describe what you want to create:", font=("Arial", 10, "bold"))
        self.desc_text = scrolledtext.ScrolledText(self.root, height=4, width=80, wrap=tk.WORD)
        self.desc_text.insert(tk.END, "Example: Create a sketch with a rectangle and extrude it to make a box")
        
        # Language selection
        self.lang_label = tk.Label(self.root, text="Programming Language:", font=("Arial", 10, "bold"))
        self.language_var = tk.StringVar(value="VBA")
        self.lang_frame = tk.Frame(self.root)
        tk.Radiobutton(self.lang_frame, text="VBA", variable=self.language_var, value="VBA").pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(self.lang_frame, text="Python", variable=self.language_var, value="Python").pack(side=tk.LEFT, padx=10)
        
        # Complexity selection
        self.complexity_label = tk.Label(self.root, text="Complexity Level:", font=("Arial", 10, "bold"))
        self.complexity_var = tk.StringVar(value="basic")
        self.complexity_combo = ttk.Combobox(
            self.root, 
            textvariable=self.complexity_var,
            values=["basic", "intermediate", "advanced"],
            state="readonly",
            width=20
        )
        
        # AI option
        self.ai_var = tk.BooleanVar()
        self.ai_check = tk.Checkbutton(
            self.root, 
            text="Use AI Generation (requires OpenAI API key)", 
            variable=self.ai_var,
            font=("Arial", 9)
        )
        
        # Buttons frame
        self.button_frame = tk.Frame(self.root)
        self.generate_btn = tk.Button(
            self.button_frame, 
            text="Generate Code", 
            command=self.generate_code,
            bg="green", 
            fg="white", 
            font=("Arial", 12, "bold"),
            width=15
        )
        self.clear_btn = tk.Button(
            self.button_frame, 
            text="Clear", 
            command=self.clear_all,
            bg="orange", 
            fg="white", 
            font=("Arial", 12),
            width=10
        )
        self.save_btn = tk.Button(
            self.button_frame, 
            text="Save Code", 
            command=self.save_code,
            bg="blue", 
            fg="white", 
            font=("Arial", 12),
            width=10
        )
        
        # Output area
        self.output_label = tk.Label(self.root, text="Generated Code:", font=("Arial", 10, "bold"))
        self.output_text = scrolledtext.ScrolledText(self.root, height=20, width=80, wrap=tk.WORD, font=("Courier", 9))
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready to generate code...")
        self.status_bar = tk.Label(
            self.root, 
            textvariable=self.status_var, 
            relief=tk.SUNKEN, 
            anchor=tk.W,
            bg="lightgray"
        )
        
        # Example buttons frame
        self.example_frame = tk.Frame(self.root)
        self.example_label = tk.Label(self.example_frame, text="Quick Examples:", font=("Arial", 9, "bold"))
        
        examples = [
            ("Create Part", "Create a new part document and set basic properties"),
            ("Draw Rectangle", "Create a sketch with a rectangle of 50x30 mm"),
            ("Extrude Feature", "Create an extrude operation with 20mm height"),
            ("Assembly", "Create an assembly and insert two parts")
        ]
        
        self.example_buttons = []
        for i, (name, desc) in enumerate(examples):
            btn = tk.Button(
                self.example_frame,
                text=name,
                command=lambda d=desc: self.load_example(d),
                bg="lightblue",
                font=("Arial", 8)
            )
            self.example_buttons.append(btn)
    
    def setup_layout(self):
        """Setup the layout of all widgets"""
        
        # Title
        self.title_label.pack(pady=10)
        
        # Description
        self.desc_label.pack(anchor=tk.W, padx=20, pady=(10, 5))
        self.desc_text.pack(padx=20, pady=5, fill=tk.X)
        
        # Language selection
        self.lang_label.pack(anchor=tk.W, padx=20, pady=(10, 5))
        self.lang_frame.pack(anchor=tk.W, padx=20)
        
        # Complexity
        self.complexity_label.pack(anchor=tk.W, padx=20, pady=(10, 5))
        self.complexity_combo.pack(anchor=tk.W, padx=20)
        
        # AI option
        self.ai_check.pack(anchor=tk.W, padx=20, pady=5)
        
        # Example buttons
        self.example_label.pack(side=tk.LEFT, padx=5)
        for btn in self.example_buttons:
            btn.pack(side=tk.LEFT, padx=2)
        self.example_frame.pack(pady=10)
        
        # Buttons
        self.generate_btn.pack(side=tk.LEFT, padx=5)
        self.clear_btn.pack(side=tk.LEFT, padx=5)
        self.save_btn.pack(side=tk.LEFT, padx=5)
        self.button_frame.pack(pady=10)
        
        # Output
        self.output_label.pack(anchor=tk.W, padx=20, pady=(10, 5))
        self.output_text.pack(padx=20, pady=5, fill=tk.BOTH, expand=True)
        
        # Status bar
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def load_example(self, description):
        """Load an example description"""
        self.desc_text.delete(1.0, tk.END)
        self.desc_text.insert(tk.END, description)
    
    def generate_code(self):
        """Generate code based on user input"""
        description = self.desc_text.get(1.0, tk.END).strip()
        
        if not description or description == "Example: Create a sketch with a rectangle and extrude it to make a box":
            messagebox.showwarning("Warning", "Please enter a description of what you want to create!")
            return
        
        # Update status
        self.status_var.set("Generating code...")
        self.root.update()
        
        try:
            # Create code request
            request = CodeRequest(
                description=description,
                language=self.language_var.get(),
                complexity=self.complexity_var.get()
            )
            
            # Generate code
            if self.ai_var.get() and self.generator.client:
                generated_code = self.generator.generate_code_with_ai(request)
                generation_method = "AI-powered"
            else:
                generated_code = self.generator.generate_template_code(request)
                generation_method = "Template-based"
            
            # Display code
            self.output_text.delete(1.0, tk.END)
            self.output_text.insert(tk.END, generated_code)
            
            # Update status
            self.status_var.set(f"Code generated successfully using {generation_method} generation!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate code: {str(e)}")
            self.status_var.set("Error occurred during code generation")
    
    def clear_all(self):
        """Clear all input and output fields"""
        self.desc_text.delete(1.0, tk.END)
        self.output_text.delete(1.0, tk.END)
        self.language_var.set("VBA")
        self.complexity_var.set("basic")
        self.ai_var.set(False)
        self.status_var.set("Ready to generate code...")
    
    def save_code(self):
        """Save the generated code to a file"""
        code = self.output_text.get(1.0, tk.END).strip()
        
        if not code:
            messagebox.showwarning("Warning", "No code to save!")
            return
        
        # Determine file extension
        ext = ".bas" if self.language_var.get() == "VBA" else ".py"
        
        # Ask user for file location
        file_path = filedialog.asksaveasfilename(
            defaultextension=ext,
            filetypes=[
                ("VBA files", "*.bas") if ext == ".bas" else ("Python files", "*.py"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            try:
                with open(file_path, 'w') as f:
                    f.write(code)
                messagebox.showinfo("Success", f"Code saved to: {file_path}")
                self.status_var.set(f"Code saved to: {os.path.basename(file_path)}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {str(e)}")

def main():
    """Main function to run the GUI application"""
    root = tk.Tk()
    app = CatiaCodeGeneratorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
