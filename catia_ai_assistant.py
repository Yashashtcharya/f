"""
CATIA V5 AI Assistant
A PyCharm application that opens CATIA V5, provides a GUI for text input,
and generates VBA code using free AI services (no API key required).
"""

import tkinter as tk
from tkinter import scrolledtext, messagebox, ttk
import win32com.client
import subprocess
import os
import requests
import json
import threading
import time

class CatiaAIAssistant:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("CATIA V5 AI Code Generator")
        self.root.geometry("800x600")
        self.catia_app = None
        self.setup_gui()
        
    def setup_gui(self):
        """Setup the main GUI interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="CATIA V5 AI Code Generator", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # CATIA Connection Section
        catia_frame = ttk.LabelFrame(main_frame, text="CATIA Connection", padding="10")
        catia_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.connect_btn = ttk.Button(catia_frame, text="Connect to CATIA V5", 
                                     command=self.connect_to_catia)
        self.connect_btn.grid(row=0, column=0, padx=(0, 10))
        
        self.status_label = ttk.Label(catia_frame, text="Status: Not Connected", 
                                     foreground="red")
        self.status_label.grid(row=0, column=1)
        
        # Input Section
        input_frame = ttk.LabelFrame(main_frame, text="Code Request", padding="10")
        input_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), 
                        pady=(0, 10))
        
        ttk.Label(input_frame, text="Describe what you want to create in CATIA:").grid(
            row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        self.input_text = scrolledtext.ScrolledText(input_frame, height=5, width=70)
        self.input_text.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Generate button
        self.generate_btn = ttk.Button(input_frame, text="Generate VBA Code", 
                                      command=self.generate_code)
        self.generate_btn.grid(row=2, column=0, pady=(0, 10))
        
        # Progress bar
        self.progress = ttk.Progressbar(input_frame, mode='indeterminate')
        self.progress.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=(0, 10))
        
        # Output Section
        output_frame = ttk.LabelFrame(main_frame, text="Generated VBA Code", padding="10")
        output_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), 
                         pady=(0, 10))
        
        self.output_text = scrolledtext.ScrolledText(output_frame, height=15, width=70)
        self.output_text.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), 
                             pady=(0, 10))
        
        # Action buttons
        button_frame = ttk.Frame(output_frame)
        button_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
        self.insert_btn = ttk.Button(button_frame, text="Insert Code to CATIA VBA", 
                                    command=self.insert_code_to_catia)
        self.insert_btn.grid(row=0, column=0, padx=(0, 10))
        
        self.copy_btn = ttk.Button(button_frame, text="Copy to Clipboard", 
                                  command=self.copy_to_clipboard)
        self.copy_btn.grid(row=0, column=1, padx=(0, 10))
        
        self.save_btn = ttk.Button(button_frame, text="Save to File", 
                                  command=self.save_to_file)
        self.save_btn.grid(row=0, column=2)
        
        # Configure grid weights for resizing
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        input_frame.columnconfigure(0, weight=1)
        output_frame.columnconfigure(0, weight=1)
        output_frame.rowconfigure(0, weight=1)
        
    def connect_to_catia(self):
        """Connect to CATIA V5 application"""
        try:
            # Try to connect to existing CATIA instance
            self.catia_app = win32com.client.Dispatch("Catia.Application")
            self.catia_app.Visible = True
            self.status_label.config(text="Status: Connected", foreground="green")
            messagebox.showinfo("Success", "Connected to CATIA V5!")
        except:
            # If no CATIA instance found, try to start it
            try:
                # You may need to adjust this path based on your CATIA installation
                catia_path = r"C:\Program Files\Dassault Systemes\B28\win_b64\code\bin\CNEXT.exe"
                if os.path.exists(catia_path):
                    subprocess.Popen([catia_path])
                    time.sleep(10)  # Wait for CATIA to start
                    self.catia_app = win32com.client.Dispatch("Catia.Application")
                    self.catia_app.Visible = True
                    self.status_label.config(text="Status: Connected", foreground="green")
                    messagebox.showinfo("Success", "Started and connected to CATIA V5!")
                else:
                    messagebox.showerror("Error", "CATIA V5 not found. Please start CATIA manually and try again.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to connect to CATIA: {str(e)}")
                
    def generate_code_thread(self, user_request):
        """Generate code in a separate thread to avoid blocking GUI"""
        try:
            self.progress.start()
            self.generate_btn.config(state='disabled')
            
            # Use HuggingFace's free inference API (no key required for some models)
            generated_code = self.generate_with_huggingface(user_request)
            
            if not generated_code:
                # Fallback to local code generation
                generated_code = self.generate_fallback_code(user_request)
            
            # Update GUI in main thread
            self.root.after(0, self.update_output, generated_code)
            
        except Exception as e:
            error_msg = f"Error generating code: {str(e)}"
            self.root.after(0, self.update_output, error_msg)
        finally:
            self.root.after(0, self.generation_complete)
    
    def generate_with_huggingface(self, user_request):
        """Generate code using HuggingFace's free inference API"""
        try:
            # Using Hugging Face's free inference endpoint
            api_url = "https://api-inference.huggingface.co/models/microsoft/DialoGPT-medium"
            
            prompt = f"""Generate VBA code for CATIA V5 based on this request: {user_request}

The code should be complete and ready to use in CATIA V5 VBA environment.
Include proper error handling and comments.

VBA Code:"""

            headers = {"Content-Type": "application/json"}
            payload = {
                "inputs": prompt,
                "parameters": {
                    "max_length": 1000,
                    "temperature": 0.7,
                    "do_sample": True
                }
            }
            
            response = requests.post(api_url, headers=headers, json=payload, timeout=30)
            
            if response.status_code == 200:
                result = response.json()
                if isinstance(result, list) and len(result) > 0:
                    generated_text = result[0].get('generated_text', '')
                    return self.clean_generated_code(generated_text, user_request)
            
        except Exception as e:
            print(f"HuggingFace API error: {e}")
            
        return None
    
    def generate_fallback_code(self, user_request):
        """Generate basic VBA code template when API fails"""
        templates = {
            "part": self.get_part_template(),
            "assembly": self.get_assembly_template(),
            "drawing": self.get_drawing_template(),
            "sketch": self.get_sketch_template(),
            "feature": self.get_feature_template()
        }
        
        request_lower = user_request.lower()
        
        # Simple keyword matching
        if any(word in request_lower for word in ["part", "solid", "extrude", "pad"]):
            return templates["part"]
        elif any(word in request_lower for word in ["assembly", "constraint", "product"]):
            return templates["assembly"]
        elif any(word in request_lower for word in ["drawing", "view", "dimension"]):
            return templates["drawing"]
        elif any(word in request_lower for word in ["sketch", "line", "circle", "profile"]):
            return templates["sketch"]
        else:
            return templates["feature"]
    
    def get_part_template(self):
        return '''Sub CreatePart()
    ' Generated VBA code for CATIA V5 Part creation
    Dim partDocument As Document
    Set partDocument = CATIA.Documents.Add("Part")
    
    Dim part As Part
    Set part = partDocument.Part
    
    Dim bodies As Bodies
    Set bodies = part.Bodies
    
    Dim body As Body
    Set body = bodies.Item("PartBody")
    
    ' Add your specific part creation code here
    ' Example: Create a sketch and extrude it
    
    ' Update the part
    part.Update
    
    MsgBox "Part created successfully!"
End Sub'''

    def get_assembly_template(self):
        return '''Sub CreateAssembly()
    ' Generated VBA code for CATIA V5 Assembly
    Dim productDocument As Document
    Set productDocument = CATIA.Documents.Add("Product")
    
    Dim product As Product
    Set product = productDocument.Product
    
    ' Add your assembly creation code here
    ' Example: Insert components and create constraints
    
    MsgBox "Assembly created successfully!"
End Sub'''

    def get_drawing_template(self):
        return '''Sub CreateDrawing()
    ' Generated VBA code for CATIA V5 Drawing
    Dim drawingDocument As Document
    Set drawingDocument = CATIA.Documents.Add("Drawing")
    
    Dim drawingRoot As DrawingRoot
    Set drawingRoot = drawingDocument.DrawingRoot
    
    ' Add your drawing creation code here
    ' Example: Create views and dimensions
    
    MsgBox "Drawing created successfully!"
End Sub'''

    def get_sketch_template(self):
        return '''Sub CreateSketch()
    ' Generated VBA code for CATIA V5 Sketch
    Dim partDocument As Document
    Set partDocument = CATIA.ActiveDocument
    
    Dim part As Part
    Set part = partDocument.Part
    
    Dim sketches As Sketches
    Set sketches = part.Bodies.Item("PartBody").Sketches
    
    ' Create a new sketch
    Dim sketch As Sketch
    Set sketch = sketches.Add(part.OriginElements.PlaneXY)
    
    ' Add your sketch geometry here
    ' Example: Create lines, circles, etc.
    
    sketch.CloseEdition
    part.Update
    
    MsgBox "Sketch created successfully!"
End Sub'''

    def get_feature_template(self):
        return '''Sub CreateFeature()
    ' Generated VBA code for CATIA V5 Feature
    Dim partDocument As Document
    Set partDocument = CATIA.ActiveDocument
    
    Dim part As Part
    Set part = partDocument.Part
    
    ' Add your feature creation code here
    ' Example: Create pads, pockets, holes, etc.
    
    part.Update
    
    MsgBox "Feature created successfully!"
End Sub'''
    
    def clean_generated_code(self, generated_text, original_request):
        """Clean and format the generated code"""
        # Extract VBA code from the response
        lines = generated_text.split('\n')
        code_lines = []
        in_code_block = False
        
        for line in lines:
            if 'Sub ' in line or 'Function ' in line:
                in_code_block = True
            if in_code_block:
                code_lines.append(line)
            if 'End Sub' in line or 'End Function' in line:
                break
                
        if code_lines:
            return '\n'.join(code_lines)
        else:
            # If no proper VBA structure found, return a template
            return self.generate_fallback_code(original_request)
    
    def generate_code(self):
        """Generate VBA code based on user input"""
        user_request = self.input_text.get("1.0", tk.END).strip()
        if not user_request:
            messagebox.showwarning("Warning", "Please enter a description of what you want to create.")
            return
            
        # Start code generation in a separate thread
        thread = threading.Thread(target=self.generate_code_thread, args=(user_request,))
        thread.daemon = True
        thread.start()
    
    def update_output(self, code):
        """Update the output text widget with generated code"""
        self.output_text.delete("1.0", tk.END)
        self.output_text.insert("1.0", code)
    
    def generation_complete(self):
        """Called when code generation is complete"""
        self.progress.stop()
        self.generate_btn.config(state='normal')
    
    def insert_code_to_catia(self):
        """Insert the generated code into CATIA VBA"""
        if not self.catia_app:
            messagebox.showerror("Error", "Please connect to CATIA first.")
            return
            
        code = self.output_text.get("1.0", tk.END).strip()
        if not code:
            messagebox.showwarning("Warning", "No code to insert.")
            return
            
        try:
            # Open VBA editor
            self.catia_app.StartCommand("VBA Editor")
            messagebox.showinfo("Info", 
                              "VBA Editor opened. Please paste the code manually or use the clipboard copy function.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open VBA editor: {str(e)}")
    
    def copy_to_clipboard(self):
        """Copy generated code to clipboard"""
        code = self.output_text.get("1.0", tk.END).strip()
        if code:
            self.root.clipboard_clear()
            self.root.clipboard_append(code)
            messagebox.showinfo("Success", "Code copied to clipboard!")
        else:
            messagebox.showwarning("Warning", "No code to copy.")
    
    def save_to_file(self):
        """Save generated code to a file"""
        code = self.output_text.get("1.0", tk.END).strip()
        if not code:
            messagebox.showwarning("Warning", "No code to save.")
            return
            
        from tkinter import filedialog
        filename = filedialog.asksaveasfilename(
            defaultextension=".bas",
            filetypes=[("VBA files", "*.bas"), ("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'w') as f:
                    f.write(code)
                messagebox.showinfo("Success", f"Code saved to {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {str(e)}")
    
    def run(self):
        """Start the application"""
        self.root.mainloop()

if __name__ == "__main__":
    app = CatiaAIAssistant()
    app.run()
