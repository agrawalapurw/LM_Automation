"""
GUI Application
User-friendly interface for email extraction and processing.
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
import os
import sys
from datetime import datetime

class EmailExtractorGUI:
    """Main GUI application."""
    
    def __init__(self, root):
        """Initialize GUI."""
        self.root = root
        self.root.title("LM Automation - Email Extractor")
        self.root.geometry("900x700")
        
        # Create notebook (tabs)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create tabs
        self.create_extract_tab()
        self.create_review_tab()
        self.create_automate_tab()
        self.create_settings_tab()
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def create_extract_tab(self):
        """Create Extract Emails tab."""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Extract Emails")
        
        # Title
        title = ttk.Label(tab, text="Extract Emails from Outlook", font=('Arial', 14, 'bold'))
        title.pack(pady=10)
        
        # Date selection
        date_frame = ttk.LabelFrame(tab, text="Date Selection", padding=10)
        date_frame.pack(fill='x', padx=20, pady=10)
        
        ttk.Label(date_frame, text="Date(s):").grid(row=0, column=0, sticky='w', pady=5)
        self.date_entry = ttk.Entry(date_frame, width=40)
        self.date_entry.grid(row=0, column=1, pady=5, padx=5)
        self.date_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))
        
        ttk.Label(date_frame, text="Format: 2024-01-15 or 2024-01-15 to 2024-01-20").grid(row=1, column=1, sticky='w')
        
        # Subject filters
        filter_frame = ttk.LabelFrame(tab, text="Subject Filters", padding=10)
        filter_frame.pack(fill='x', padx=20, pady=10)
        
        self.use_default_filters = tk.BooleanVar(value=True)
        ttk.Checkbutton(filter_frame, text="Use default filters (Pre-MQL ready for review/validation)", 
                       variable=self.use_default_filters).pack(anchor='w')
        
        # Move emails option
        move_frame = ttk.LabelFrame(tab, text="Email Moving", padding=10)
        move_frame.pack(fill='x', padx=20, pady=10)
        
        self.move_emails = tk.BooleanVar(value=False)
        ttk.Checkbutton(move_frame, text="Move emails to distribution partner folders during extraction", 
                       variable=self.move_emails).pack(anchor='w')
        
        # Extract button
        btn_frame = ttk.Frame(tab)
        btn_frame.pack(pady=20)
        
        self.extract_btn = ttk.Button(btn_frame, text="Extract Emails", command=self.extract_emails, width=20)
        self.extract_btn.pack()
        
        # Log output
        log_frame = ttk.LabelFrame(tab, text="Log", padding=10)
        log_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        self.extract_log = scrolledtext.ScrolledText(log_frame, height=15, state='disabled')
        self.extract_log.pack(fill='both', expand=True)
    
    def create_review_tab(self):
        """Create Review & Move tab."""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Review & Move")
        
        title = ttk.Label(tab, text="Review Excel and Move Emails", font=('Arial', 14, 'bold'))
        title.pack(pady=10)
        
        # File selection
        file_frame = ttk.LabelFrame(tab, text="Select Excel File", padding=10)
        file_frame.pack(fill='x', padx=20, pady=10)
        
        self.excel_file_var = tk.StringVar()
        ttk.Label(file_frame, textvariable=self.excel_file_var).pack(side='left', padx=5)
        ttk.Button(file_frame, text="Browse...", command=self.browse_excel).pack(side='right')
        
        # Actions
        action_frame = ttk.LabelFrame(tab, text="Actions", padding=10)
        action_frame.pack(fill='x', padx=20, pady=10)
        
        ttk.Button(action_frame, text="Open Excel for Review", command=self.open_excel, width=25).pack(pady=5)
        ttk.Button(action_frame, text="Move Emails to Folders", command=self.move_emails_action, width=25).pack(pady=5)
        
        # Log
        log_frame = ttk.LabelFrame(tab, text="Log", padding=10)
        log_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        self.review_log = scrolledtext.ScrolledText(log_frame, height=15, state='disabled')
        self.review_log.pack(fill='both', expand=True)
    
    def create_automate_tab(self):
        """Create Automate Forms tab."""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Automate Forms")
        
        title = ttk.Label(tab, text="Automate Web Form Submission", font=('Arial', 14, 'bold'))
        title.pack(pady=10)
        
        # File selection
        file_frame = ttk.LabelFrame(tab, text="Select Excel File", padding=10)
        file_frame.pack(fill='x', padx=20, pady=10)
        
        self.automate_file_var = tk.StringVar()
        ttk.Label(file_frame, textvariable=self.automate_file_var).pack(side='left', padx=5)
        ttk.Button(file_frame, text="Browse...", command=self.browse_excel_automate).pack(side='right')
        
        # Options
        options_frame = ttk.LabelFrame(tab, text="Options", padding=10)
        options_frame.pack(fill='x', padx=20, pady=10)
        
        ttk.Label(options_frame, text="Sheet to process:").grid(row=0, column=0, sticky='w', pady=5)
        self.sheet_choice = ttk.Combobox(options_frame, values=["Validation", "Review", "Both"], state='readonly', width=20)
        self.sheet_choice.grid(row=0, column=1, pady=5, padx=5)
        self.sheet_choice.current(2)
        
        self.headless_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="Run in headless mode (no browser UI)", 
                       variable=self.headless_var).grid(row=1, column=0, columnspan=2, sticky='w', pady=5)
        
        # Start button
        btn_frame = ttk.Frame(tab)
        btn_frame.pack(pady=20)
        
        self.automate_btn = ttk.Button(btn_frame, text="Start Automation", command=self.start_automation, width=20)
        self.automate_btn.pack()
        
        # Log
        log_frame = ttk.LabelFrame(tab, text="Log", padding=10)
        log_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        self.automate_log = scrolledtext.ScrolledText(log_frame, height=15, state='disabled')
        self.automate_log.pack(fill='both', expand=True)
    
    def create_settings_tab(self):
        """Create Settings tab."""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Settings")
        
        title = ttk.Label(tab, text="Settings", font=('Arial', 14, 'bold'))
        title.pack(pady=10)
        
        # Validation data
        validation_frame = ttk.LabelFrame(tab, text="Validation Data", padding=10)
        validation_frame.pack(fill='x', padx=20, pady=10)
        
        ttk.Label(validation_frame, text="Validation data folder:").grid(row=0, column=0, sticky='w', pady=5)
        self.validation_folder_var = tk.StringVar(value="validation_data")
        ttk.Entry(validation_frame, textvariable=self.validation_folder_var, width=40).grid(row=0, column=1, pady=5, padx=5)
        ttk.Button(validation_frame, text="Browse...", command=self.browse_validation_folder).grid(row=0, column=2, padx=5)
        
        ttk.Button(validation_frame, text="Refresh Validation Data", command=self.refresh_validation_data, width=25).grid(row=1, column=1, pady=10)
        
        # About
        about_frame = ttk.LabelFrame(tab, text="About", padding=10)
        about_frame.pack(fill='x', padx=20, pady=10)
        
        about_text = """LM Automation Tool v1.0
        
Email extraction and processing for Pre-MQL leads.

For help and documentation, see USER_GUIDE.md"""
        
        ttk.Label(about_frame, text=about_text, justify='left').pack(anchor='w')
    
    def log_message(self, log_widget, message):
        """Add message to log widget."""
        log_widget.config(state='normal')
        log_widget.insert(tk.END, message + '\n')
        log_widget.see(tk.END)
        log_widget.config(state='disabled')
        self.root.update()
    
    def extract_emails(self):
        """Extract emails (runs main.py logic)."""
        self.extract_btn.config(state='disabled')
        self.status_var.set("Extracting emails...")
        self.extract_log.config(state='normal')
        self.extract_log.delete(1.0, tk.END)
        self.extract_log.config(state='disabled')
        
        def run():
            try:
                self.log_message(self.extract_log, "Starting email extraction...")
                self.log_message(self.extract_log, "This feature requires command-line interaction.")
                self.log_message(self.extract_log, "Please use main.py directly for now.")
                self.log_message(self.extract_log, "\nTo run: python main.py")
                messagebox.showinfo("Info", "Please run main.py from command line for email extraction.")
            except Exception as e:
                self.log_message(self.extract_log, f"Error: {e}")
                messagebox.showerror("Error", str(e))
            finally:
                self.extract_btn.config(state='normal')
                self.status_var.set("Ready")
        
        thread = threading.Thread(target=run)
        thread.start()
    
    def browse_excel(self):
        """Browse for Excel file."""
        output_dir = os.path.join(os.path.dirname(__file__), "output")
        filename = filedialog.askopenfilename(
            initialdir=output_dir,
            title="Select Excel file",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )
        if filename:
            self.excel_file_var.set(filename)
    
    def browse_excel_automate(self):
        """Browse for Excel file for automation."""
        output_dir = os.path.join(os.path.dirname(__file__), "output")
        filename = filedialog.askopenfilename(
            initialdir=output_dir,
            title="Select Excel file",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )
        if filename:
            self.automate_file_var.set(filename)
    
    def browse_validation_folder(self):
        """Browse for validation data folder."""
        folder = filedialog.askdirectory(title="Select validation data folder")
        if folder:
            self.validation_folder_var.set(folder)
    
    def open_excel(self):
        """Open Excel file."""
        excel_path = self.excel_file_var.get()
        if not excel_path:
            messagebox.showwarning("Warning", "Please select an Excel file first.")
            return
        
        if not os.path.exists(excel_path):
            messagebox.showerror("Error", "File not found.")
            return
        
        os.startfile(excel_path)
        self.log_message(self.review_log, f"Opened: {excel_path}")
    
    def move_emails_action(self):
        """Move emails based on Excel."""
        excel_path = self.excel_file_var.get()
        if not excel_path:
            messagebox.showwarning("Warning", "Please select an Excel file first.")
            return
        
        self.log_message(self.review_log, "Email moving requires command-line interaction.")
        self.log_message(self.review_log, "Please use move_emails.py directly.")
        self.log_message(self.review_log, "\nTo run: python move_emails.py")
        messagebox.showinfo("Info", "Please run move_emails.py from command line.")
    
    def start_automation(self):
        """Start form automation."""
        excel_path = self.automate_file_var.get()
        if not excel_path:
            messagebox.showwarning("Warning", "Please select an Excel file first.")
            return
        
        self.log_message(self.automate_log, "Form automation requires command-line interaction.")
        self.log_message(self.automate_log, "Please use automate_forms.py directly.")
        self.log_message(self.automate_log, "\nTo run: python automate_forms.py")
        messagebox.showinfo("Info", "Please run automate_forms.py from command line.")
    
    def refresh_validation_data(self):
        """Refresh validation data."""
        messagebox.showinfo("Info", "Validation data will be reloaded on next extraction.")
        self.log_message(self.extract_log, "Validation data will be refreshed on next run.")

def main():
    """Main entry point."""
    root = tk.Tk()
    app = EmailExtractorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()