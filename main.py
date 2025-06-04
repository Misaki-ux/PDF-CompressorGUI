import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import subprocess
import os
import shutil
import threading # To keep the GUI responsive during compression
import os.path

# --- Ghostscript Compression Logic (from previous example) ---
def compress_pdf_ghostscript(input_path, output_path, quality_preset='ebook', status_callback=None):
    """
    Compresses a PDF using Ghostscript (external tool).
    Offers good compression, especially for images.
    Requires Ghostscript to be installed.
    status_callback: A function to send status updates to.
    """
    if status_callback:
        status_callback(f"Starting compression for: {os.path.basename(input_path)}")

    if not os.path.exists(input_path):
        message = f"Error: Input file '{input_path}' not found."
        if status_callback: status_callback(message)
        messagebox.showerror("Error", message)
        return False, 0, 0

    # Check common Ghostscript installation paths
    possible_paths = [
        r"C:\Program Files\gs\gs10.05.1\bin\gswin64c.exe",  # 64-bit specific version
        r"C:\Program Files\gs\gs10.05.1\bin\gswin32c.exe",  # 32-bit specific version
        r"C:\Program Files\gs\gs*\bin\gswin64c.exe",        # Any 64-bit version
        r"C:\Program Files\gs\gs*\bin\gswin32c.exe",        # Any 32-bit version
    ]

    gs_executable = None
    # First try the PATH
    gs_executable = shutil.which("gs") or shutil.which("gswin64c") or shutil.which("gswin32c")
    
    # If not found in PATH, try common installation paths
    if not gs_executable:
        for path in possible_paths:
            # Handle wildcard paths
            if '*' in path:
                import glob
                matches = glob.glob(path)
                if matches:
                    gs_executable = matches[-1]  # Use the last match (usually highest version)
                    break
            elif os.path.exists(path):
                gs_executable = path
                break

    if not gs_executable:
        message = ("Error: Ghostscript not found in PATH or common installation directories.\n"
                  "Please ensure Ghostscript is installed and either:\n"
                  "1. Add its bin directory to your PATH, or\n"
                  "2. Install it in the standard location (C:\\Program Files\\gs\\).\n"
                  "Download from: https://www.ghostscript.com/releases/gsdnld.html")
        if status_callback: status_callback(message)
        messagebox.showerror("Ghostscript Not Found", message)
        return False, 0, 0

    if status_callback:
        status_callback(f"Found Ghostscript at: {gs_executable}")

    original_size = os.path.getsize(input_path)

    try:
        args = [
            gs_executable,
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            f"-dPDFSETTINGS=/{quality_preset}",
            "-dNOPAUSE",
            "-dQUIET",
            "-dBATCH",
            f"-sOutputFile={output_path}",
            input_path
        ]
        if status_callback:
            status_callback(f"Running Ghostscript with preset: {quality_preset}...")
            status_callback(f"Command: {' '.join(args)}")


        # Run Ghostscript
        process = subprocess.Popen(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
        stdout, stderr = process.communicate() # Wait for it to finish

        if process.returncode != 0:
            error_message = f"Ghostscript failed (return code {process.returncode})."
            if stderr:
                error_message += f"\nError Output:\n{stderr}"
            if stdout: # Sometimes errors also go to stdout
                error_message += f"\nStandard Output:\n{stdout}"

            if status_callback: status_callback(error_message)
            messagebox.showerror("Ghostscript Error", error_message)
            # Clean up potentially partially created output file
            if os.path.exists(output_path) and os.path.getsize(output_path) == 0:
                os.remove(output_path)
            return False, original_size, 0


        if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
            message = f"Ghostscript ran, but output file '{output_path}' is missing or empty."
            if stderr: message += f"\nGS Stderr: {stderr}"
            if stdout: message += f"\nGS Stdout: {stdout}"
            if status_callback: status_callback(message)
            messagebox.showwarning("Compression Issue", message)
            return False, original_size, 0

        compressed_size = os.path.getsize(output_path)
        message = f"Compression complete! Output: {os.path.basename(output_path)}"
        if status_callback: status_callback(message)
        return True, original_size, compressed_size

    except FileNotFoundError: # Should be caught by gs_executable check, but good to have
        message = "Error: Ghostscript command not found. Is it installed and in PATH?"
        if status_callback: status_callback(message)
        messagebox.showerror("Error", message)
        return False, original_size, 0
    except Exception as e:
        message = f"An unexpected error occurred: {e}"
        if status_callback: status_callback(message)
        messagebox.showerror("Error", message)
        return False, original_size, 0

# --- Tkinter GUI Application ---
class PDFCompressorApp:
    def __init__(self, root):
        self.root = root
        root.title("PDF Compressor")
        root.geometry("600x500")
        
        # Configure dark theme colors
        self.colors = {
            'bg': '#2b2b2b',
            'fg': '#ffffff',
            'button': '#3c3f41',
            'button_fg': '#ffffff',
            'entry': '#333333',         # Darker grey for entry fields
            'entry_fg': '#ffffff',      # White text for entries
            'highlight': '#4CAF50',
            'text_bg': '#1e1e1e'
        }
        
        # Apply dark theme to root
        root.configure(bg=self.colors['bg'])
        style = ttk.Style()
        
        # Configure frame style
        style.configure('Dark.TFrame', background=self.colors['bg'])
        style.configure('Dark.TLabel', foreground=self.colors['fg'], background=self.colors['bg'])
        
        # Configure the Dark.TEntry style with dark background and white text
        style.configure('Dark.TEntry', 
            fieldbackground=self.colors['entry'],
            foreground=self.colors['entry_fg'],
            insertcolor=self.colors['entry_fg'],
            selectbackground='#4a6984',
            selectforeground='#ffffff'
        )
        
        # Override the Entry widget settings
        style.layout('Dark.TEntry', [
            ('Entry.plain.field', {'children': [
                ('Entry.background', {'children': [
                    ('Entry.padding', {'children': [
                        ('Entry.textarea', {'sticky': 'nswe'})
                    ], 'sticky': 'nswe'})
                ], 'sticky': 'nswe'})
            ], 'border': '1', 'sticky': 'nswe'})
        ])
        
        # Map states for the entry
        style.map('Dark.TEntry',
            fieldbackground=[
                ('readonly', self.colors['entry']),
                ('disabled', self.colors['entry']),
                ('active', self.colors['entry'])
            ],
            foreground=[
                ('readonly', self.colors['entry_fg']),
                ('disabled', self.colors['entry_fg']),
                ('active', self.colors['entry_fg'])
            ],
            selectbackground=[('readonly', '#4a6984')],
            selectforeground=[('readonly', '#ffffff')]
        )
        
        style.configure('Dark.TButton', background=self.colors['button'], foreground=self.colors['button_fg'])
        
        # Configure styles for OptionMenu (menu button part)
        style.layout('Dark.TMenubutton', [
            ('Menubutton.background', {
                'children': [
                    ('Menubutton.button', {
                        'children': [
                            ('Menubutton.focus', {
                                'children': [
                                    ('Menubutton.padding', {
                                        'children': [
                                            ('Menubutton.label', {'side': 'left', 'expand': 1})
                                        ],
                                        'expand': 1,
                                        'sticky': 'nswe'
                                    })
                                ],
                                'sticky': 'nswe'
                            })
                        ],
                        'sticky': 'nswe'
                    })
                ],
                'sticky': 'nswe'
            })
        ])
        
        style.configure('Dark.TMenubutton',
            background=self.colors['entry'],
            foreground=self.colors['entry_fg'],
            selectbackground='#4a6984',
            selectforeground='#ffffff',
            borderwidth=1,
            relief='raised'
        )
        
        style.map('Dark.TMenubutton',
            background=[('active', self.colors['entry'])],
            foreground=[('active', self.colors['entry_fg'])]
        )
        
        # Configure styles for OptionMenu
        style.configure('Dark.TOptionMenu', 
            background=self.colors['entry'],
            foreground=self.colors['entry_fg'],
            fieldbackground=self.colors['entry'],
            selectbackground='#4a6984',
            selectforeground='#ffffff'
        )
        
        style.map('Dark.TOptionMenu',
            background=[('active', self.colors['entry'])],
            fieldbackground=[('readonly', self.colors['entry'])],
            foreground=[('active', self.colors['entry_fg'])]
        )
        
        self.input_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        self.quality_preset = tk.StringVar(value='ebook') # Default preset

        # --- UI Elements ---

        # Input File Selection
        ttk.Label(root, text="Input PDF File:", style='Dark.TLabel').grid(row=0, column=0, padx=5, pady=5, sticky="w")
        input_frame = ttk.Frame(root, style='Dark.TFrame')  # Add style here
        input_frame.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        self.input_entry = ttk.Entry(
        input_frame, 
        textvariable=self.input_file_path, 
        state="readonly", 
        style='Dark.TEntry'
        )
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
    
        
        browse_btn = tk.Button(input_frame, text="Browse...", command=self.browse_input_file,
                             bg=self.colors['button'], fg=self.colors['button_fg'],
                             activebackground=self.colors['highlight'])
        browse_btn.pack(side=tk.RIGHT)

        # Output File Selection
        ttk.Label(root, text="Output PDF File:", style='Dark.TLabel').grid(row=1, column=0, padx=5, pady=5, sticky="w")
        output_frame = ttk.Frame(root, style='Dark.TFrame')  # Add style here
        output_frame.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        
        self.output_entry = ttk.Entry(
        output_frame, 
        textvariable=self.output_file_path, 
        state="readonly", 
        style='Dark.TEntry'
        )
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        # Quality Preset
        ttk.Label(root, text="Quality Preset:", style='Dark.TLabel').grid(row=2, column=0, padx=5, pady=5, sticky="w")
        presets = ['screen', 'ebook', 'printer', 'prepress', 'default']
        self.preset_menu = ttk.OptionMenu(root, self.quality_preset, self.quality_preset.get(), *presets, style='Dark.TMenubutton')
        self.preset_menu.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        # Configure the dropdown menu colors
        menu = self.preset_menu['menu']
        menu.configure(
            bg=self.colors['entry'],
            fg=self.colors['entry_fg'],
            activebackground=self.colors['button'],
            activeforeground=self.colors['button_fg'],
            relief='flat',
            borderwidth=0
        )

        # Compress Button
        self.compress_button = tk.Button(root, text="Compress PDF",
                                       command=self.start_compression_thread,
                                       bg=self.colors['highlight'],
                                       fg=self.colors['button_fg'],
                                       font=("Arial", 12, "bold"),
                                       activebackground=self.colors['button'])
        self.compress_button.grid(row=3, column=0, columnspan=3, padx=10, pady=15, ipady=5, sticky="ew")

        # Status Area
        ttk.Label(root, text="Status:", style='Dark.TLabel').grid(row=4, column=0, padx=5, pady=5, sticky="nw")
        self.status_text = tk.Text(root, height=10, wrap=tk.WORD,
                                 bg=self.colors['text_bg'],
                                 fg=self.colors['fg'],
                                 relief=tk.SUNKEN,
                                 borderwidth=1)
        self.status_text.grid(row=5, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")

        # Configure grid column weights for responsiveness
        root.grid_columnconfigure(1, weight=1)
        root.grid_rowconfigure(5, weight=1) # Allow status text area to expand

        self.add_status("Ready. Please select input file...")

    def add_status(self, message):
        self.status_text.config(state="normal")
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END) # Scroll to the bottom
        self.status_text.config(state="disabled")
        self.root.update_idletasks() # Ensure GUI updates immediately

    def browse_input_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Input PDF",
            filetypes=(("PDF files", "*.pdf"), ("All files", "*.*"))
        )
        if file_path:
            self.input_file_path.set(file_path)
            # Automatically set output path with _compressed suffix
            base, ext = os.path.splitext(file_path)
            output_path = f"{base}_compressed{ext}"
            self.output_file_path.set(output_path)
            self.add_status(f"Input file selected: {os.path.basename(file_path)}")
            self.add_status(f"Output will be saved as: {os.path.basename(output_path)}")

    def browse_output_file(self):
        # Suggest a filename if input is selected
        initial_file = ""
        if self.input_file_path.get():
            base, ext = os.path.splitext(self.input_file_path.get())
            initial_file = f"{os.path.basename(base)}_compressed{ext}"
        
        file_path = filedialog.asksaveasfilename(
            title="Save Compressed PDF As",
            defaultextension=".pdf",
            initialfile=initial_file,
            filetypes=(("PDF files", "*.pdf"), ("All files", "*.*"))
        )
        if file_path:
            self.output_file_path.set(file_path)
            self.add_status(f"Output file set to: {file_path}")

    def start_compression_thread(self):
        input_p = self.input_file_path.get()
        output_p = self.output_file_path.get()
        quality = self.quality_preset.get()

        if not input_p:
            messagebox.showerror("Error", "Please select an input PDF file.")
            self.add_status("Error: No input file selected.")
            return
        if not output_p:
            messagebox.showerror("Error", "Please specify an output PDF file.")
            self.add_status("Error: No output file specified.")
            return
        
        if input_p == output_p:
            messagebox.showerror("Error", "Input and output file paths cannot be the same. This would overwrite your original file.")
            self.add_status("Error: Input and output paths are identical.")
            return

        self.compress_button.config(state="disabled", text="Compressing...")
        self.add_status("--- Starting Compression ---")
        
        # Run compression in a separate thread to keep GUI responsive
        thread = threading.Thread(target=self.run_compression, args=(input_p, output_p, quality))
        thread.daemon = True # Allows main program to exit even if thread is running
        thread.start()

    def run_compression(self, input_p, output_p, quality):
        success, original_size, compressed_size = compress_pdf_ghostscript(
            input_p, output_p, quality, self.add_status
        )

        if success:
            self.add_status(f"Original size: {original_size / 1024:.2f} KB")
            self.add_status(f"Compressed size: {compressed_size / 1024:.2f} KB")
            if original_size > 0:
                reduction = (1 - compressed_size / original_size) * 100
                self.add_status(f"Reduction: {reduction:.2f}%")
            self.add_status(f"Compressed PDF saved as: {output_p}")
            messagebox.showinfo("Success", f"PDF compressed successfully!\nSaved as: {output_p}")
        else:
            self.add_status("Compression failed or was cancelled.")
            # Error messages are already shown by compress_pdf_ghostscript or here

        # Re-enable button on the main thread
        self.root.after(0, self.enable_compress_button)

    def enable_compress_button(self):
        self.compress_button.config(state="normal", text="Compress PDF")


if __name__ == "__main__":
    # Use the same path detection logic as in the compression function
    possible_paths = [
        r"C:\Program Files\gs\gs10.05.1\bin\gswin64c.exe",
        r"C:\Program Files\gs\gs10.05.1\bin\gswin32c.exe",
        r"C:\Program Files\gs\gs*\bin\gswin64c.exe",
        r"C:\Program Files\gs\gs*\bin\gswin32c.exe",
    ]

    gs_executable_check = shutil.which("gs") or shutil.which("gswin64c") or shutil.which("gswin32c")
    
    if not gs_executable_check:
        for path in possible_paths:
            if '*' in path:
                import glob
                matches = glob.glob(path)
                if matches:
                    gs_executable_check = matches[-1]
                    break
            elif os.path.exists(path):
                gs_executable_check = path
                break

    if not gs_executable_check:
        temp_root = tk.Tk()
        temp_root.withdraw()
        messagebox.showwarning(
            "Ghostscript Missing",
            "Ghostscript was not found on your system.\n"
            "Please ensure it's installed in C:\\Program Files\\gs\\ or add it to your PATH.\n"
            "Download from: https://www.ghostscript.com/releases/gsdnld.html"
        )
        temp_root.destroy()

    main_root = tk.Tk()
    
    # Configure the style after creating the main window
    style = ttk.Style(main_root)
    style.configure('Dark.TMessage',
                   background='#2b2b2b',
                   foreground='#ffffff')
    
    app = PDFCompressorApp(main_root)
    main_root.mainloop()