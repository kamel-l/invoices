import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from tkinter.font import Font
import shutil
from datetime import datetime

class FileRenamerGUI:
    def __init__(self, root):
        self.root = root
        self.setup_window()
        self.setup_variables()
        self.setup_styles()
        self.create_widgets()
        self.files_to_rename = []
        
    def setup_window(self):
        """Configure main window"""
        self.root.title("🔧 File Renamer - Replace %20 with -")
        self.root.geometry("900x700")
        self.root.minsize(850, 650)
        
        # Center window
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (900 // 2)
        y = (self.root.winfo_screenheight() // 2) - (700 // 2)
        self.root.geometry(f"900x700+{x}+{y}")
        
        self.root.configure(bg='#f0f2f5')
        
    def setup_variables(self):
        """Initialize variables"""
        self.folder_path = tk.StringVar()
        self.find_text = tk.StringVar(value='%20')
        self.replace_text = tk.StringVar(value='-')
        self.file_count = tk.StringVar(value="0 files found")
        self.preview_mode = tk.BooleanVar(value=True)
        
    def setup_styles(self):
        """Configure ttk styles"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Primary button style
        style.configure('Primary.TButton',
                       background='#4CAF50',
                       foreground='white',
                       font=('Segoe UI', 10, 'bold'),
                       padding=(20, 10))
        style.map('Primary.TButton',
                 background=[('active', '#45a049'), ('pressed', '#3d8b40')])
        
        # Secondary button style
        style.configure('Secondary.TButton',
                       background='#2196F3',
                       foreground='white',
                       font=('Segoe UI', 9),
                       padding=(15, 8))
        style.map('Secondary.TButton',
                 background=[('active', '#1976D2'), ('pressed', '#1565C0')])
        
        # Danger button style
        style.configure('Danger.TButton',
                       background='#f44336',
                       foreground='white',
                       font=('Segoe UI', 9),
                       padding=(15, 8))
        style.map('Danger.TButton',
                 background=[('active', '#da190b'), ('pressed', '#c41408')])
        
    def create_widgets(self):
        """Create user interface"""
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # Header
        self.create_header(main_frame)
        
        # Folder selection section
        self.create_folder_section(main_frame)
        
        # Find and replace section
        self.create_find_replace_section(main_frame)
        
        # Preview section
        self.create_preview_section(main_frame)
        
        # Action buttons
        self.create_action_buttons(main_frame)
        
    def create_header(self, parent):
        """Create header"""
        header_frame = ttk.Frame(parent)
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        
        title_font = Font(family='Segoe UI', size=24, weight='bold')
        title_label = ttk.Label(header_frame, text="File Renamer Tool", 
                               font=title_font, foreground='#4CAF50')
        title_label.grid(row=0, column=0, sticky=tk.W)
        
        subtitle_font = Font(family='Segoe UI', size=12)
        subtitle_label = ttk.Label(header_frame, text="Replace text in multiple file names", 
                                  font=subtitle_font, foreground='#666666')
        subtitle_label.grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        
    def create_folder_section(self, parent):
        """Folder selection section"""
        folder_frame = ttk.LabelFrame(parent, text="📁 Folder Selection", padding="15")
        folder_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        folder_frame.columnconfigure(1, weight=1)
        
        ttk.Label(folder_frame, text="Folder:").grid(row=0, column=0, sticky=tk.W, pady=5)
        
        path_frame = ttk.Frame(folder_frame)
        path_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=5)
        path_frame.columnconfigure(0, weight=1)
        
        self.folder_entry = ttk.Entry(path_frame, textvariable=self.folder_path, 
                                     width=60, state='readonly')
        self.folder_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(path_frame, text="Browse", 
                  command=self.browse_folder).grid(row=0, column=1)
        
        # File count label
        self.count_label = ttk.Label(folder_frame, textvariable=self.file_count, 
                                    foreground='#666666')
        self.count_label.grid(row=1, column=1, sticky=tk.W, padx=(10, 0), pady=(5, 0))
        
    def create_find_replace_section(self, parent):
        """Find and replace section"""
        replace_frame = ttk.LabelFrame(parent, text="🔍 Find and Replace", padding="15")
        replace_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        replace_frame.columnconfigure(1, weight=1)
        
        # Find text
        ttk.Label(replace_frame, text="Find:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.find_entry = ttk.Entry(replace_frame, textvariable=self.find_text, width=40)
        self.find_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=5)
        
        # Replace text
        ttk.Label(replace_frame, text="Replace with:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.replace_entry = ttk.Entry(replace_frame, textvariable=self.replace_text, width=40)
        self.replace_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=5)
        
        # Scan button
        ttk.Button(replace_frame, text="🔎 Scan Files", 
                  style='Secondary.TButton',
                  command=self.scan_files).grid(row=2, column=1, sticky=tk.E, pady=(10, 0))
        
    def create_preview_section(self, parent):
        """Preview section"""
        preview_frame = ttk.LabelFrame(parent, text="👁️ Preview Changes", padding="15")
        preview_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        
        # Create treeview for file preview
        columns = ('original', 'new', 'status')
        self.tree = ttk.Treeview(preview_frame, columns=columns, show='headings', height=15)
        
        # Define headings
        self.tree.heading('original', text='Original Name')
        self.tree.heading('new', text='New Name')
        self.tree.heading('status', text='Status')
        
        # Define column widths
        self.tree.column('original', width=350)
        self.tree.column('new', width=350)
        self.tree.column('status', width=100)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
    def create_action_buttons(self, parent):
        """Action buttons"""
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=4, column=0, pady=(10, 0))
        
        self.rename_btn = ttk.Button(button_frame, text="✅ Rename Files", 
                                    style='Primary.TButton',
                                    command=self.rename_files,
                                    state='disabled')
        self.rename_btn.grid(row=0, column=0, padx=(0, 15))
        
        ttk.Button(button_frame, text="🔄 Reset", 
                  style='Secondary.TButton',
                  command=self.reset).grid(row=0, column=1, padx=(0, 15))
        
        ttk.Button(button_frame, text="🗑️ Clear Preview", 
                  command=self.clear_preview).grid(row=0, column=2)
        
    def browse_folder(self):
        """Browse for folder"""
        folder = filedialog.askdirectory(title="Select Folder")
        if folder:
            self.folder_path.set(folder)
            # Auto-scan after selecting folder
            self.scan_files()
            
    def scan_files(self):
        """Scan files in selected folder"""
        folder = self.folder_path.get()
        if not folder:
            messagebox.showwarning("Warning", "Please select a folder first")
            return
            
        if not os.path.exists(folder):
            messagebox.showerror("Error", "Selected folder does not exist")
            return
            
        find_text = self.find_text.get()
        if not find_text:
            messagebox.showwarning("Warning", "Please enter text to find")
            return
            
        # Clear previous results
        self.clear_preview()
        self.files_to_rename = []
        
        # Scan files
        try:
            all_files = os.listdir(folder)
            matching_files = [f for f in all_files if find_text in f and os.path.isfile(os.path.join(folder, f))]
            
            if not matching_files:
                messagebox.showinfo("Info", f"No files found containing '{find_text}'")
                self.file_count.set("0 files found")
                return
            
            # Add to preview
            replace_text = self.replace_text.get()
            for filename in matching_files:
                new_name = filename.replace(find_text, replace_text)
                self.files_to_rename.append((filename, new_name))
                self.tree.insert('', tk.END, values=(filename, new_name, 'Ready'))
            
            self.file_count.set(f"{len(matching_files)} files found")
            self.rename_btn.config(state='normal')
            
            messagebox.showinfo("Success", f"Found {len(matching_files)} files to rename")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error scanning folder: {str(e)}")
            
    def rename_files(self):
        """Rename files"""
        if not self.files_to_rename:
            messagebox.showwarning("Warning", "No files to rename")
            return
            
        # Confirm action
        result = messagebox.askyesno("Confirm", 
                                     f"Are you sure you want to rename {len(self.files_to_rename)} files?\n\n"
                                     "This action cannot be undone!")
        if not result:
            return
            
        folder = self.folder_path.get()
        success_count = 0
        error_count = 0
        
        # Clear tree
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Rename files
        for old_name, new_name in self.files_to_rename:
            try:
                old_path = os.path.join(folder, old_name)
                new_path = os.path.join(folder, new_name)
                
                # Check if new file already exists
                if os.path.exists(new_path):
                    self.tree.insert('', tk.END, values=(old_name, new_name, '⚠️ Exists'))
                    error_count += 1
                    continue
                
                # Rename file
                os.rename(old_path, new_path)
                self.tree.insert('', tk.END, values=(old_name, new_name, '✅ Success'))
                success_count += 1
                
            except Exception as e:
                self.tree.insert('', tk.END, values=(old_name, new_name, f'❌ Error'))
                error_count += 1
        
        # Show summary
        messagebox.showinfo("Complete", 
                           f"Renaming complete!\n\n"
                           f"✅ Success: {success_count}\n"
                           f"❌ Failed: {error_count}")
        
        self.files_to_rename = []
        self.rename_btn.config(state='disabled')
        
    def clear_preview(self):
        """Clear preview tree"""
        for item in self.tree.get_children():
            self.tree.delete(item)
            
    def reset(self):
        """Reset everything"""
        self.clear_preview()
        self.files_to_rename = []
        self.folder_path.set("")
        self.file_count.set("0 files found")
        self.rename_btn.config(state='disabled')

def main():
    """Main function"""
    root = tk.Tk()
    app = FileRenamerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()