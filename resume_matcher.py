import PyPDF2
import re
import numpy as np
from sentence_transformers import SentenceTransformer
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
from tkinter.font import Font

class BulkResumeMatcher:
    def __init__(self):
        self.model = SentenceTransformer('all-MiniLM-L6-v2')
        self.resume_files = []
        
        # Color scheme
        self.bg_color = "#f0f2f5"
        self.primary_color = "#4f46e5"  # Indigo
        self.secondary_color = "#10b981"  # Emerald
        self.accent_color = "#f59e0b"  # Amber
        self.error_color = "#ef4444"  # Red
        self.text_color = "#1f2937"
        self.light_text = "#6b7280"
        
        # Main window setup
        self.root = tk.Tk()
        self.root.title("Bulk Resume Matcher Pro")
        self.root.geometry("1100x750")
        self.root.configure(bg=self.bg_color)
        
        # Custom fonts
        self.title_font = Font(family="Segoe UI", size=14, weight="bold")
        self.button_font = Font(family="Segoe UI", size=10, weight="bold")
        
        # Configure styles
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('TFrame', background=self.bg_color)
        self.style.configure('TLabel', background=self.bg_color, foreground=self.text_color)
        self.style.configure('TButton', font=self.button_font, background=self.primary_color, 
                           foreground="white", borderwidth=0)
        self.style.map('TButton',
            background=[('active', '#4338ca'), ('disabled', '#a5b4fc')],
            foreground=[('disabled', '#d1d5db')]
        )
        
        # Create UI
        self.create_widgets()
        
    def create_widgets(self):
        # Main container
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Left panel - Inputs
        input_frame = ttk.Frame(main_frame, style='TFrame')
        input_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        # Job Description Section
        ttk.Label(input_frame, text="Job Description:", font=self.title_font, 
                 foreground=self.primary_color).pack(anchor='w')
        
        self.jd_text = scrolledtext.ScrolledText(input_frame, wrap=tk.WORD,
                                               height=15, font=('Segoe UI', 10),
                                               padx=10, pady=10, relief="solid",
                                               borderwidth=1, bg="white",
                                               fg=self.text_color)
        self.jd_text.pack(fill=tk.BOTH, expand=True)
        
        # Resume Upload Section
        upload_frame = ttk.Frame(input_frame)
        upload_frame.pack(fill=tk.X, pady=10)
        
        upload_btn = ttk.Button(upload_frame, text="📂 Upload Multiple Resumes", 
                              command=self.upload_multiple_resumes,
                              style='TButton')
        upload_btn.pack(side=tk.LEFT)
        
        self.status_label = ttk.Label(upload_frame, text="No resumes loaded",
                                    foreground=self.light_text)
        self.status_label.pack(side=tk.LEFT, padx=10)
        
        # Right panel - Results
        results_frame = ttk.Frame(main_frame)
        results_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        ttk.Label(results_frame, text="Matching Results:", font=self.title_font,
                foreground=self.primary_color).pack(anchor='w')
        
        # Treeview for results with custom styling
        self.style.configure("Treeview", 
                            background="white",
                            foreground=self.text_color,
                            rowheight=25,
                            fieldbackground="white")
        self.style.map('Treeview', background=[('selected', '#e5e7eb')])
        
        self.results_tree = ttk.Treeview(results_frame, columns=('Score', 'Missing'), 
                                       selectmode='extended', height=25)
        self.results_tree.heading('#0', text='Resume File', anchor='w')
        self.results_tree.heading('Score', text='Match Score', anchor='center')
        self.results_tree.heading('Missing', text='Missing Keywords', anchor='w')
        
        self.results_tree.column('#0', width=250, anchor='w')
        self.results_tree.column('Score', width=100, anchor='center')
        self.results_tree.column('Missing', width=300, anchor='w')
        
        scrollbar = ttk.Scrollbar(results_frame, orient="vertical", 
                                command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=scrollbar.set)
        
        self.results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Analyze button with accent color
        self.style.configure('Accent.TButton', background=self.secondary_color)
        analyze_btn = ttk.Button(main_frame, text="🔍 Analyze All Resumes", 
                               command=self.analyze_all_resumes,
                               style='Accent.TButton')
        analyze_btn.pack(pady=10)
    
    def upload_multiple_resumes(self):
        filepaths = filedialog.askopenfilenames(
            title="Select Resume PDFs",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        
        if filepaths:
            self.resume_files = list(filepaths)
            self.status_label.config(text=f"{len(self.resume_files)} resumes loaded")
            messagebox.showinfo("Success", f"Successfully loaded {len(self.resume_files)} resumes",
                              parent=self.root)
    
    def extract_text(self, filepath):
        with open(filepath, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            text = " ".join([page.extract_text() for page in reader.pages if page.extract_text()])
            return re.sub(r'\s+', ' ', text).strip()
    
    def analyze_all_resumes(self):
        jd = self.jd_text.get("1.0", tk.END).strip()
        
        if not jd:
            messagebox.showwarning("Warning", "Please enter a job description",
                                 parent=self.root)
            return
            
        if not self.resume_files:
            messagebox.showwarning("Warning", "Please upload at least one resume",
                                 parent=self.root)
            return
        
        try:
            # Clear previous results
            for item in self.results_tree.get_children():
                self.results_tree.delete(item)
            
            # Encode job description once
            jd_embedding = self.model.encode(jd)
            jd_words = set(re.findall(r'\b\w{3,}\b', jd.lower()))
            
            results = []
            
            # Process each resume
            for filepath in self.resume_files:
                try:
                    filename = filepath.split("/")[-1]
                    resume_text = self.extract_text(filepath)
                    resume_embedding = self.model.encode(resume_text)
                    similarity = np.dot(jd_embedding, resume_embedding)
                    score = round(similarity * 100, 1)
                    
                    # Find missing keywords
                    resume_words = set(re.findall(r'\b\w{3,}\b', resume_text.lower()))
                    missing_keywords = sorted(jd_words - resume_words)[:5]
                    missing_str = ", ".join(missing_keywords) if missing_keywords else "None"
                    
                    results.append((filename, score, missing_str))
                    
                except Exception as e:
                    results.append((filename, "Error", str(e)))
            
            # Sort by score (highest first)
            results.sort(key=lambda x: x[1] if isinstance(x[1], (int, float)) else -1, reverse=True)
            
            # Add to treeview with color coding
            for filename, score, missing in results:
                if isinstance(score, (int, float)):
                    if score >= 75:
                        tag = "high"
                    elif score >= 50:
                        tag = "medium"
                    else:
                        tag = "low"
                else:
                    tag = "error"
                
                self.results_tree.insert('', 'end', text=filename, 
                                       values=(f"{score}%" if isinstance(score, (int, float)) else score, missing),
                                       tags=(tag,))
            
            # Configure tag colors
            self.results_tree.tag_configure('high', foreground=self.secondary_color)
            self.results_tree.tag_configure('medium', foreground=self.accent_color)
            self.results_tree.tag_configure('low', foreground=self.error_color)
            self.results_tree.tag_configure('error', foreground=self.error_color)
            
            messagebox.showinfo("Analysis Complete", f"Processed {len(results)} resumes",
                              parent=self.root)
            
        except Exception as e:
            messagebox.showerror("Error", f"Analysis failed: {str(e)}",
                               parent=self.root)

if __name__ == "__main__":
    try:
        import sentence_transformers, PyPDF2, numpy
    except ImportError:
        import subprocess
        import sys
        subprocess.run([sys.executable, "-m", "pip", "install", 
                       "sentence-transformers", "PyPDF2", "numpy"])
    
    app = BulkResumeMatcher()
    app.root.mainloop()
