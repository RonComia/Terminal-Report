import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import os
import sys
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from PIL import Image, ImageTk
import glob
import re
import threading

class TreeInventoryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Tree Inventory Manager")
        self.root.configure(bg="#1E1E1E")
        
        self.root.geometry("800x600")
        self.root.minsize(600, 400)
        
        self.main_frame = tk.Frame(root, bg="#1E1E1E")
        self.main_frame.place(relx=0.05, rely=0.05, relwidth=0.9, relheight=0.9)
        
        self.create_header()
        self.create_content()
        
        self.excel_file_path = None
        self.image_folder_path = None
        self.output_folder_path = None
        self.start_row = 2
        self.batch_size = 1002
        self.num_batches_to_process = None
        
        # Add these new attributes
        self.image_pattern = "auto_detect"  # Default to auto-detect
        self.use_subfolders = True
        
        self.image_cache = {}
        self.pattern_stats = {}

    def create_header(self):
        header_frame = tk.Frame(self.main_frame, bg="#1E1E1E")
        header_frame.place(relx=0.5, rely=0.05, relwidth=0.9, relheight=0.1, anchor="n")
        
        self.inventory_btn = tk.Button(
            header_frame, 
            text="INVENTORY", 
            bg="#1E1E1E", 
            fg="#FFC107",
            activebackground="#2A2A2A",
            activeforeground="#FFC107",
            bd=0,
            font=("Arial", 10, "bold"),
            command=self.show_inventory
        )
        self.inventory_btn.pack(side=tk.LEFT, padx=10)
        
        self.cutting_btn = tk.Button(
            header_frame, 
            text="CUTTING", 
            bg="#1E1E1E", 
            fg="#696969",
            activebackground="#2A2A2A",
            activeforeground="#FFC107",
            bd=0,
            font=("Arial", 10, "bold"),
            state=tk.DISABLED
        )
        self.cutting_btn.pack(side=tk.LEFT, padx=10)

    def create_content(self):
        content_frame = tk.Frame(self.main_frame, bg="#1E1E1E")
        content_frame.place(relx=0.5, rely=0.2, relwidth=0.8, relheight=0.7, anchor="n")
        
        self.excel_btn = tk.Button(
            content_frame,
            text="INSERT EXCEL FILE",
            bg="#1E1E1E",
            fg="#FFFFFF",
            activebackground="#2A2A2A",
            activeforeground="#FFFFFF",
            bd=1,
            relief=tk.SOLID,
            font=("Arial", 10, "bold"),
            width=20,
            height=2,
            command=self.select_excel_file
        )
        self.excel_btn.place(relx=0.5, rely=0.5, anchor="center")

    def show_inventory(self):
        self.clear_content()
        
        inventory_frame = tk.Frame(self.main_frame, bg="#1E1E1E")
        inventory_frame.place(relx=0.5, rely=0.2, relwidth=0.8, relheight=0.7, anchor="n")
        
        select_excel_btn = tk.Button(
            inventory_frame,
            text="Select Excel File",
            bg="#1E1E1E",
            fg="#FFFFFF",
            activebackground="#2A2A2A",
            activeforeground="#FFFFFF",
            bd=1,
            relief=tk.SOLID,
            font=("Arial", 10),
            command=self.select_excel_file
        )
        select_excel_btn.pack(pady=5)
        
        select_images_btn = tk.Button(
            inventory_frame,
            text="Select Images Folder",
            bg="#1E1E1E",
            fg="#FFFFFF",
            activebackground="#2A2A2A",
            activeforeground="#FFFFFF",
            bd=1,
            relief=tk.SOLID,
            font=("Arial", 10),
            command=self.select_image_folder
        )
        select_images_btn.pack(pady=5)
        
        select_output_btn = tk.Button(
            inventory_frame,
            text="Select Output Folder",
            bg="#1E1E1E",
            fg="#FFFFFF",
            activebackground="#2A2A2A",
            activeforeground="#FFFFFF",
            bd=1,
            relief=tk.SOLID,
            font=("Arial", 10),
            command=self.select_output_folder
        )
        select_output_btn.pack(pady=5)
        
        # Add settings button
        settings_btn = tk.Button(
            inventory_frame,
            text="Image Naming Settings",
            bg="#1E1E1E",
            fg="#FFFFFF",
            activebackground="#2A2A2A",
            activeforeground="#FFFFFF",
            bd=1,
            relief=tk.SOLID,
            font=("Arial", 10),
            command=self.show_image_settings
        )
        settings_btn.pack(pady=5)
        
        config_frame = tk.Frame(inventory_frame, bg="#1E1E1E", bd=1, relief=tk.SOLID)
        config_frame.pack(pady=10, padx=20, fill=tk.X)
        
        start_row_frame = tk.Frame(config_frame, bg="#1E1E1E")
        start_row_frame.pack(pady=5, fill=tk.X)
        
        start_row_label = tk.Label(
            start_row_frame,
            text="Start Row:",
            bg="#1E1E1E",
            fg="#FFFFFF",
            font=("Arial", 10),
            width=15,
            anchor="e"
        )
        start_row_label.pack(side=tk.LEFT, padx=5)
        
        self.start_row_var = tk.StringVar(value=str(self.start_row))
        start_row_entry = tk.Entry(
            start_row_frame,
            textvariable=self.start_row_var,
            bg="#2A2A2A",
            fg="#FFFFFF",
            insertbackground="#FFFFFF",
            width=10
        )
        start_row_entry.pack(side=tk.LEFT, padx=5)
        
        batch_size_frame = tk.Frame(config_frame, bg="#1E1E1E")
        batch_size_frame.pack(pady=5, fill=tk.X)
        
        batch_size_label = tk.Label(
            batch_size_frame,
            text="Batch Size:",
            bg="#1E1E1E",
            fg="#FFFFFF",
            font=("Arial", 10),
            width=15,
            anchor="e"
        )
        batch_size_label.pack(side=tk.LEFT, padx=5)
        
        self.batch_size_var = tk.StringVar(value=str(self.batch_size))
        batch_size_entry = tk.Entry(
            batch_size_frame,
            textvariable=self.batch_size_var,
            bg="#2A2A2A",
            fg="#FFFFFF",
            insertbackground="#FFFFFF",
            width=10
        )
        batch_size_entry.pack(side=tk.LEFT, padx=5)
        
        num_batches_frame = tk.Frame(config_frame, bg="#1E1E1E")
        num_batches_frame.pack(pady=5, fill=tk.X)
        
        num_batches_label = tk.Label(
            num_batches_frame,
            text="Batches to Process:",
            bg="#1E1E1E",
            fg="#FFFFFF",
            font=("Arial", 10),
            width=15,
            anchor="e"
        )
        num_batches_label.pack(side=tk.LEFT, padx=5)
        
        self.num_batches_var = tk.StringVar(value="All")
        num_batches_entry = tk.Entry(
            num_batches_frame,
            textvariable=self.num_batches_var,
            bg="#2A2A2A",
            fg="#FFFFFF",
            insertbackground="#FFFFFF",
            width=10
        )
        num_batches_entry.pack(side=tk.LEFT, padx=5)
        
        help_label = tk.Label(
            num_batches_frame,
            text="(Enter 'All' or a number)",
            bg="#1E1E1E",
            fg="#888888",
            font=("Arial", 8)
        )
        help_label.pack(side=tk.LEFT, padx=5)
        
        self.process_btn = tk.Button(
            inventory_frame,
            text="Generate Word Documents",
            bg="#1E1E1E",
            fg="#FFFFFF",
            activebackground="#2A2A2A",
            activeforeground="#FFFFFF",
            bd=1,
            relief=tk.SOLID,
            font=("Arial", 10, "bold"),
            command=self.process_data,
            state=tk.DISABLED
        )
        self.process_btn.pack(pady=10)
        
        self.status_frame = tk.Frame(inventory_frame, bg="#1E1E1E")
        self.status_frame.pack(fill=tk.X, pady=10)
        
        self.excel_label = tk.Label(
            self.status_frame,
            text="Excel File: Not selected",
            bg="#1E1E1E",
            fg="#888888",
            font=("Arial", 8),
            anchor="w"
        )
        self.excel_label.pack(fill=tk.X)
        
        self.images_label = tk.Label(
            self.status_frame,
            text="Images Folder: Not selected",
            bg="#1E1E1E",
            fg="#888888",
            font=("Arial", 8),
            anchor="w"
        )
        self.images_label.pack(fill=tk.X)
        
        self.output_label = tk.Label(
            self.status_frame,
            text="Output Folder: Not selected",
            bg="#1E1E1E",
            fg="#888888",
            font=("Arial", 8),
            anchor="w"
        )
        self.output_label.pack(fill=tk.X)
        
        # Add settings status label
        self.settings_frame = tk.Frame(self.status_frame, bg="#1E1E1E")
        self.settings_frame.pack(fill=tk.X)
        
        pattern_name = next((name for name, val in [
            ("Simple numbers", "simple_numbers"),
            ("Tree prefix", "tree_prefix"),
            ("Alpha-numeric", "alpha_numeric"),
            ("Number with suffix", "number_suffix"),
            ("Auto-detect", "auto_detect")
        ] if val == self.image_pattern), "Unknown")
        
        self.settings_label = tk.Label(
            self.settings_frame,
            text=f"Pattern: {pattern_name} | Subfolders: {'Yes' if self.use_subfolders else 'No'}",
            bg="#1E1E1E",
            fg="#888888",
            font=("Arial", 8),
            anchor="w"
        )
        self.settings_label.pack(fill=tk.X)
        
        self.progress_frame = tk.Frame(inventory_frame, bg="#1E1E1E")
        self.progress_frame.pack(fill=tk.X, pady=10)
        
        self.progress_label = tk.Label(
            self.progress_frame,
            text="Progress:",
            bg="#1E1E1E",
            fg="#FFFFFF",
            font=("Arial", 8),
        )
        self.progress_label.pack(anchor="w")
        
        self.progress_bar = ttk.Progressbar(
            self.progress_frame,
            orient="horizontal",
            length=300,
            mode="determinate"
        )
        self.progress_bar.pack(fill=tk.X, pady=5)
        
        self.progress_frame.pack_forget()

    def clear_content(self):
        for widget in self.main_frame.winfo_children():
            if widget != self.main_frame.winfo_children()[0]:
                widget.destroy()

    def select_excel_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            self.excel_file_path = file_path
            if hasattr(self, 'excel_label'):
                self.excel_label.config(
                    text=f"Excel File: {os.path.basename(file_path)}",
                    fg="#FFFFFF"
                )
            else:
                self.show_inventory()
                self.excel_label.config(
                    text=f"Excel File: {os.path.basename(file_path)}",
                    fg="#FFFFFF"
                )
            self.check_all_selected()

    def select_image_folder(self):
        folder_path = filedialog.askdirectory(
            title="Select Folder Containing Tree Images"
        )
        
        if folder_path:
            self.image_folder_path = folder_path
            self.images_label.config(
                text=f"Images Folder: {os.path.basename(folder_path)}",
                fg="#FFFFFF"
            )
            self.check_all_selected()

    def select_output_folder(self):
        folder_path = filedialog.askdirectory(
            title="Select Output Folder for Word Documents"
        )
        
        if folder_path:
            self.output_folder_path = folder_path
            self.output_label.config(
                text=f"Output Folder: {os.path.basename(folder_path)}",
                fg="#FFFFFF"
            )
            self.check_all_selected()

    def check_all_selected(self):
        if (self.excel_file_path and self.image_folder_path and self.output_folder_path):
            self.process_btn.config(state=tk.NORMAL)
        else:
            self.process_btn.config(state=tk.DISABLED)

    def show_image_settings(self):
        """Display a dialog to configure image naming pattern"""
        settings_window = tk.Toplevel(self.root)
        settings_window.title("Image Naming Settings")
        settings_window.geometry("400x300")
        settings_window.transient(self.root)
        settings_window.grab_set()
        
        settings_window.geometry("+%d+%d" % (
            self.root.winfo_rootx() + self.root.winfo_width() // 2 - 200,
            self.root.winfo_rooty() + self.root.winfo_height() // 2 - 150
        ))
        
        # Frame for settings
        frame = tk.Frame(settings_window, bg="#1E1E1E", padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Create pattern selection
        pattern_label = tk.Label(
            frame, 
            text="Select image naming pattern:",
            bg="#1E1E1E",
            fg="#FFFFFF",
            font=("Arial", 10, "bold"),
            anchor="w"
        )
        pattern_label.pack(fill=tk.X, pady=(0, 10))
        
        # Pattern options
        self.pattern_var = tk.StringVar(value=self.image_pattern)
        
        patterns = [
            ("Simple numbers (e.g., '123.jpg')", "simple_numbers"),
            ("Tree prefix (e.g., 'T123.jpg', 'Tree123.jpg')", "tree_prefix"),
            ("Alpha-numeric (e.g., 'A123.jpg', 'B001.jpg')", "alpha_numeric"),
            ("Number with suffix (e.g., '123A.jpg', '001B.jpg')", "number_suffix"),
            ("Auto-detect (slower but handles mixed formats)", "auto_detect")
        ]
        
        for text, pattern in patterns:
            rb = tk.Radiobutton(
                frame,
                text=text,
                variable=self.pattern_var,
                value=pattern,
                bg="#1E1E1E",
                fg="#FFFFFF",
                selectcolor="#2A2A2A",
                activebackground="#2A2A2A",
                activeforeground="#FFFFFF"
            )
            rb.pack(anchor="w", pady=2)
        
        # Additional options
        options_frame = tk.Frame(frame, bg="#1E1E1E", pady=10)
        options_frame.pack(fill=tk.X)
        
        self.use_subfolders_var = tk.BooleanVar(value=self.use_subfolders)
        subfolder_cb = tk.Checkbutton(
            options_frame,
            text="Search in subfolders (slower)",
            variable=self.use_subfolders_var,
            bg="#1E1E1E",
            fg="#FFFFFF",
            selectcolor="#2A2A2A",
            activebackground="#2A2A2A",
            activeforeground="#FFFFFF"
        )
        subfolder_cb.pack(anchor="w")
        
        # Buttons
        button_frame = tk.Frame(frame, bg="#1E1E1E", pady=10)
        button_frame.pack(fill=tk.X)
        
        save_btn = tk.Button(
            button_frame,
            text="Save Settings",
            bg="#1E1E1E",
            fg="#FFFFFF",
            activebackground="#2A2A2A",
            activeforeground="#FFFFFF",
            bd=1,
            relief=tk.SOLID,
            font=("Arial", 10),
            command=lambda: self.save_image_settings(settings_window)
        )
        save_btn.pack(side=tk.RIGHT, padx=5)
        
        cancel_btn = tk.Button(
            button_frame,
            text="Cancel",
            bg="#1E1E1E",
            fg="#FFFFFF",
            activebackground="#2A2A2A",
            activeforeground="#FFFFFF",
            bd=1,
            relief=tk.SOLID,
            font=("Arial", 10),
            command=settings_window.destroy
        )
        cancel_btn.pack(side=tk.RIGHT, padx=5)

    def save_image_settings(self, window):
        """Save the selected image naming pattern"""
        self.image_pattern = self.pattern_var.get()
        self.use_subfolders = self.use_subfolders_var.get()
        
        # Clear existing cache
        self.image_cache = {}
        
        # Close the window
        window.destroy()
        
        # Update status label
        if hasattr(self, 'settings_label'):
            pattern_name = next((name for name, val in [
                ("Simple numbers", "simple_numbers"),
                ("Tree prefix", "tree_prefix"),
                ("Alpha-numeric", "alpha_numeric"),
                ("Number with suffix", "number_suffix"),
                ("Auto-detect", "auto_detect")
            ] if val == self.image_pattern), "Unknown")
            
            self.settings_label.config(
                text=f"Pattern: {pattern_name} | Subfolders: {'Yes' if self.use_subfolders else 'No'}",
                fg="#FFFFFF"
            )

    def process_data(self):
        """Optimized data processing method"""
        try:
            # --- Input validation ---
            try:
                self.start_row = int(self.start_row_var.get())
                if self.start_row < 2:
                    raise ValueError("Start row must be 2 or greater (row 1 is header)")
            except ValueError:
                messagebox.showerror("Invalid Input", "Start row must be a valid number (2 or greater)")
                return
            
            try:
                self.batch_size = int(self.batch_size_var.get())
                if self.batch_size < 1:
                    raise ValueError("Batch size must be at least 1")
            except ValueError:
                messagebox.showerror("Invalid Input", "Batch size must be a valid number (1 or greater)")
                return
            
            if self.num_batches_var.get().strip().lower() == "all":
                self.num_batches_to_process = None
            else:
                try:
                    self.num_batches_to_process = int(self.num_batches_var.get())
                    if self.num_batches_to_process < 1:
                        raise ValueError("Number of batches must be at least 1")
                except ValueError:
                    messagebox.showerror("Invalid Input", "Batches to process must be 'All' or a valid number")
                    return
            
            # Show progress UI
            self.progress_frame.pack(fill=tk.X, pady=10)
            self.progress_bar["value"] = 0
            self.root.update()
            
            # --- Load and prepare Excel data ---
            # Only read the needed columns to save memory
            try:
                self.progress_label.config(text="Loading Excel data...")
                self.root.update()
                
                # Use usecols parameter to only load required columns
                df = pd.read_excel(
                    self.excel_file_path, 
                    usecols=['TREE NUMBER', 'SPECIES'],
                    dtype={'TREE NUMBER': str, 'SPECIES': str}  # Specify dtypes to avoid conversion
                )
            except ValueError:
                try:
                    df = pd.read_excel(
                        self.excel_file_path, 
                        usecols=['Tree Number', 'Species'],
                        dtype={'Tree Number': str, 'Species': str}
                    )
                    df.rename(columns={'Tree Number': 'TREE NUMBER', 'Species': 'SPECIES'}, inplace=True)
                except ValueError:
                    # Load excel first, then figure out the columns
                    df = pd.read_excel(self.excel_file_path)
                    
                    # Handle column renames more efficiently
                    columns_to_check = [
                        ('TREE NUMBER', 'Tree Number', df.columns[5] if len(df.columns) >= 6 else None),
                        ('SPECIES', 'Species', df.columns[6] if len(df.columns) >= 7 else None)
                    ]
                    
                    for target_col, alt_col, fallback_col in columns_to_check:
                        if target_col in df.columns:
                            pass  # Already has the correct name
                        elif alt_col in df.columns:
                            df.rename(columns={alt_col: target_col}, inplace=True)
                        elif fallback_col is not None:
                            df.rename(columns={fallback_col: target_col}, inplace=True)
                        else:
                            messagebox.showerror(
                                "Column Error", 
                                f"Could not find required column: {target_col}"
                            )
                            self.progress_frame.pack_forget()
                            return
            
            # Extract only needed columns to save memory
            df = df[['TREE NUMBER', 'SPECIES']].copy()
            
            # Skip rows if needed
            if self.start_row > 2:
                start_index = self.start_row - 2
                df = df.iloc[start_index:].reset_index(drop=True)
            
            # --- Load image cache in background ---
            self.progress_label.config(text="Loading images...")
            self.root.update()
            
            # Load image cache
            self.load_image_cache()
            
            # --- Process in batches with memory management ---
            total_rows = len(df)
            num_batches = (total_rows + self.batch_size - 1) // self.batch_size
            
            if self.num_batches_to_process is not None:
                num_batches = min(num_batches, self.num_batches_to_process)
            
            self.progress_bar["maximum"] = num_batches
            
            # Initialize new caches
            if not hasattr(self, 'match_cache'):
                self.match_cache = {}
            if not hasattr(self, 'image_size_cache'):
                self.image_size_cache = {}
            
            for batch_num in range(num_batches):
                self.progress_label.config(text=f"Processing batch {batch_num + 1} of {num_batches}...")
                self.progress_bar["value"] = batch_num
                self.root.update()
                
                batch_start = batch_num * self.batch_size
                batch_end = min((batch_num + 1) * self.batch_size, total_rows)
                
                # Create a copy to avoid modifying the original dataframe
                batch_df = df.iloc[batch_start:batch_end].copy()
                
                # Generate Word document
                self.generate_word_doc(batch_df, batch_num + 1, num_batches)
                
                # Update progress
                self.progress_bar["value"] = batch_num + 1
                self.root.update()
                
                # Clean memory periodically
                if batch_num % 5 == 0:
                    # Clear caches to free memory
                    self.match_cache.clear()
                    if len(self.image_size_cache) > 100:  # Only clear if it's getting large
                        self.image_size_cache.clear()
                    
                    # Force garbage collection
                    import gc
                    gc.collect()
            
            # Clean up
            if hasattr(self, 'match_cache'):
                self.match_cache.clear()
            if hasattr(self, 'image_size_cache'):
                self.image_size_cache.clear()
                
            # Remove temp directory if it exists
            temp_dir = os.path.join(self.output_folder_path, 'temp_images')
            if os.path.exists(temp_dir):
                try:
                    import shutil
                    shutil.rmtree(temp_dir)
                except:
                    pass
            
            self.progress_frame.pack_forget()
            messagebox.showinfo(
                "Processing Complete", 
                f"Generated {num_batches} Word documents in {self.output_folder_path}"
            )
            
        except Exception as e:
            self.progress_frame.pack_forget()
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            import traceback
            traceback.print_exc()

    def load_image_cache(self):
        """Optimized image cache loader with better performance"""
        self.image_cache = {}
        
        image_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tif', '.tiff']
        all_files = []
        
        # Use a set for faster lookups
        image_extensions_set = set(image_extensions)
        
        # Collect files first
        if self.use_subfolders:
            for root, dirs, files in os.walk(self.image_folder_path):
                for file in files:
                    ext = os.path.splitext(file.lower())[1]
                    if ext in image_extensions_set:
                        all_files.append(os.path.join(root, file))
        else:
            for file in os.listdir(self.image_folder_path):
                ext = os.path.splitext(file.lower())[1]
                if ext in image_extensions_set:
                    all_files.append(os.path.join(self.image_folder_path, file))
        
        total_files = len(all_files)
        if total_files > 0:
            progress_window = tk.Toplevel(self.root)
            progress_window.title("Loading Images")
            progress_window.geometry("300x80")
            progress_window.transient(self.root)
            progress_window.grab_set()
            
            progress_window.geometry("+%d+%d" % (
                self.root.winfo_rootx() + self.root.winfo_width() // 2 - 150,
                self.root.winfo_rooty() + self.root.winfo_height() // 2 - 40
            ))
            
            progress_label = tk.Label(
                progress_window,
                text=f"Indexing image files (0/{total_files})...",
                font=("Arial", 10)
            )
            progress_label.pack(pady=5)
            
            image_progress = ttk.Progressbar(
                progress_window,
                orient="horizontal",
                length=280,
                mode="determinate",
                maximum=total_files
            )
            image_progress.pack(pady=5, padx=10)
            
            pattern_type = self.image_pattern
            
            # Compile regex patterns only once
            patterns = {}
            if pattern_type == "simple_numbers" or pattern_type == "auto_detect":
                patterns['num'] = re.compile(r'(?:^|[^a-zA-Z0-9])(\d+)(?:[^a-zA-Z0-9]|$)')
            if pattern_type == "tree_prefix" or pattern_type == "auto_detect":
                patterns['tree'] = re.compile(r'(?:^|[^a-zA-Z])([Tt](?:ree)?)[- _]?(\d+)')
            if pattern_type == "alpha_numeric" or pattern_type == "auto_detect":
                patterns['alpha'] = re.compile(r'([a-zA-Z]+)[- _]?(\d+)')
            if pattern_type == "number_suffix" or pattern_type == "auto_detect":
                patterns['num_suffix'] = re.compile(r'(\d+)[- _]?([a-zA-Z]+)')
            if pattern_type == "auto_detect":
                patterns['all_nums'] = re.compile(r'\d+')
            
            # Process in larger batches for better performance
            batch_size = 200
            for i in range(0, total_files, batch_size):
                batch_end = min(i + batch_size, total_files)
                current_batch = all_files[i:batch_end]
                
                progress_label.config(text=f"Indexing image files ({i+1}/{total_files})...")
                image_progress["value"] = i + 1
                progress_window.update()
                
                for file_path in current_batch:
                    filename = os.path.basename(file_path)
                    name_without_ext = os.path.splitext(filename)[0]
                    
                    # Always add the exact filename match for quick lookup
                    self.image_cache[f"file_{name_without_ext.lower()}"] = file_path
                    
                    # Process filename with appropriate pattern based on selected mode
                    if pattern_type == "simple_numbers" or pattern_type == "auto_detect":
                        self._process_simple_numbers(file_path, name_without_ext, patterns)
                    elif pattern_type == "tree_prefix":
                        self._process_tree_prefix(file_path, name_without_ext, patterns)
                    elif pattern_type == "alpha_numeric":
                        self._process_alpha_numeric(file_path, name_without_ext, patterns)
                    elif pattern_type == "number_suffix":
                        self._process_number_suffix(file_path, name_without_ext, patterns)
                
                image_progress["value"] = batch_end
                progress_window.update()
            
            progress_window.destroy()

    def _process_simple_numbers(self, file_path, name, patterns):
        matches = patterns['num'].finditer(name)
        for match in matches:
            num_str = match.group(1)
            try:
                num_val = int(num_str)
                self.image_cache[f"num_{num_val}"] = file_path
                # Add zero-padded numbers without zeros
                if len(num_str) > 1 and num_str.startswith('0'):
                    self.image_cache[f"raw_{num_val}"] = file_path
            except ValueError:
                pass
    
    def _process_tree_prefix(self, file_path, name, patterns):
        matches = patterns['tree'].finditer(name)
        for match in matches:
            try:
                num_val = int(match.group(2))
                self.image_cache[f"tree_{num_val}"] = file_path
            except ValueError:
                pass
    
    def _process_alpha_numeric(self, file_path, name, patterns):
        matches = patterns['alpha'].finditer(name)
        for match in matches:
            prefix = match.group(1).lower()
            try:
                num_val = int(match.group(2))
                self.image_cache[f"alpha_{prefix}_{num_val}"] = file_path
            except ValueError:
                pass
    
    def _process_number_suffix(self, file_path, name, patterns):
        matches = patterns['num_suffix'].finditer(name)
        for match in matches:
            try:
                num_val = int(match.group(1))
                suffix = match.group(2).lower()
                self.image_cache[f"numsuf_{num_val}_{suffix}"] = file_path
                # Also add without suffix as fallback
                self.image_cache[f"numsuf_{num_val}"] = file_path
            except ValueError:
                pass

    def get_image_path(self, tree_number):
        """Optimized image path lookup based on the selected pattern"""
        if not tree_number or pd.isna(tree_number):
            return None
        
        try:
            tree_num_str = str(tree_number).strip()
            
            # Quick lookup for previously matched patterns (new cache)
            if hasattr(self, 'match_cache') and tree_num_str in self.match_cache:
                return self.match_cache[tree_num_str]
            
            if not hasattr(self, 'match_cache'):
                self.match_cache = {}
                
            pattern_type = self.image_pattern
            result = None
            
            # Simple direct lookups first (fastest)
            tree_num_lower = tree_num_str.lower()
            if f"file_{tree_num_lower}" in self.image_cache:
                result = self.image_cache[f"file_{tree_num_lower}"]
            
            # Try numeric matching if no direct match
            if not result:
                try:
                    numeric_tree_num = int(tree_num_str)
                    
                    # Try most common pattern first based on selected type
                    if pattern_type == "simple_numbers" or pattern_type == "auto_detect":
                        if f"num_{numeric_tree_num}" in self.image_cache:
                            result = self.image_cache[f"num_{numeric_tree_num}"]
                        elif f"raw_{numeric_tree_num}" in self.image_cache:
                            result = self.image_cache[f"raw_{numeric_tree_num}"]
                    
                    if not result and (pattern_type == "tree_prefix" or pattern_type == "auto_detect"):
                        if f"tree_{numeric_tree_num}" in self.image_cache:
                            result = self.image_cache[f"tree_{numeric_tree_num}"]
                            
                    if not result and pattern_type == "auto_detect":
                        if f"any_{numeric_tree_num}" in self.image_cache:
                            result = self.image_cache[f"any_{numeric_tree_num}"]
                        elif f"numsuf_{numeric_tree_num}" in self.image_cache:
                            result = self.image_cache[f"numsuf_{numeric_tree_num}"]
                except ValueError:
                    pass
            
            # Only do expensive regex operations if simpler methods failed
            if not result:
                if pattern_type in ["alpha_numeric", "auto_detect"]:
                    match = re.match(r'^([a-zA-Z]+)[- _]?(\d+)$', tree_num_str)
                    if match:
                        prefix = match.group(1).lower()
                        try:
                            num_val = int(match.group(2))
                            if f"alpha_{prefix}_{num_val}" in self.image_cache:
                                result = self.image_cache[f"alpha_{prefix}_{num_val}"]
                        except ValueError:
                            pass
                
                if not result and pattern_type in ["number_suffix", "auto_detect"]:
                    match = re.match(r'^(\d+)[- _]?([a-zA-Z]+)$', tree_num_str)
                    if match:
                        try:
                            num_val = int(match.group(1))
                            suffix = match.group(2).lower()
                            if f"numsuf_{num_val}_{suffix}" in self.image_cache:
                                result = self.image_cache[f"numsuf_{num_val}_{suffix}"]
                        except ValueError:
                            pass
            
            # Cache the result for this tree number
            self.match_cache[tree_num_str] = result
            return result
            
        except Exception as e:
            print(f"Error matching tree {tree_number}: {str(e)}")
            return None

    def optimize_image_for_word(self, image_path, max_width=1024, max_height=768):
        """Resize image to reduce Word document size and improve performance"""
        try:
            if not hasattr(self, 'image_size_cache'):
                self.image_size_cache = {}
                
            # Check if we've already processed this image
            if image_path in self.image_size_cache:
                return self.image_size_cache[image_path]
                
            with Image.open(image_path) as img:
                # Check if resizing is needed
                if img.width > max_width or img.height > max_height:
                    # Calculate new dimensions while maintaining aspect ratio
                    ratio = min(max_width/img.width, max_height/img.height)
                    new_width = int(img.width * ratio)
                    new_height = int(img.height * ratio)
                    
                    # Create temp path for optimized image
                    temp_dir = os.path.join(self.output_folder_path, 'temp_images')
                    if not os.path.exists(temp_dir):
                        os.makedirs(temp_dir)
                        
                    file_name = os.path.basename(image_path)
                    temp_path = os.path.join(temp_dir, f"opt_{file_name}")
                    
                    # Only create optimized version if it doesn't exist
                    if not os.path.exists(temp_path):
                        img_resized = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                        img_resized.save(temp_path, quality=85, optimize=True)
                    
                    self.image_size_cache[image_path] = temp_path
                    return temp_path
                else:
                    # Image already smaller than max dimensions
                    self.image_size_cache[image_path] = image_path
                    return image_path
                    
        except Exception as e:
            print(f"Error optimizing image {image_path}: {str(e)}")
            return image_path  # Return original path on error

    def generate_word_doc(self, batch_df, batch_num, total_batches):
        """Optimized Word document generation"""
        doc = Document()
        
        for section in doc.sections:
            section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Inches(0.5)
        
        actual_rows = len(batch_df)
        trees_per_column = (actual_rows + 1) // 2
        table_rows = trees_per_column * 2
        
        table = doc.add_table(rows=table_rows, cols=2)
        table.style = 'Table Grid'
        
        # Set column widths once
        for col in table.columns:
            for cell in col.cells:
                cell.width = Inches(3.25)
        
        # Set row heights once
        for i in range(table_rows):
            table.rows[i].height = Inches(2.5 if i % 2 == 0 else 0.4)
        
        # Prepare font settings - define once
        arial_11_bold = {'name': 'Arial', 'size': Pt(11), 'bold': True}
        
        # Process images in smaller groups to reduce memory usage
        chunk_size = 10  # Process 10 trees at a time
        for chunk_start in range(0, actual_rows, chunk_size):
            chunk_end = min(chunk_start + chunk_size, actual_rows)
            
            for tree_index in range(chunk_start, chunk_end):
                tree_row = batch_df.iloc[tree_index]
                tree_number = str(tree_row['TREE NUMBER'])
                species = str(tree_row['SPECIES']).upper()
                
                # Calculate cell position
                row_index = tree_index // 2 * 2
                col_index = tree_index % 2
                
                # Image cell
                image_cell = table.cell(row_index, col_index)
                image_para = image_cell.paragraphs[0]
                image_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Get and optimize image
                image_path = self.get_image_path(tree_number)
                
                if image_path and os.path.exists(image_path):
                    try:
                        # Optimize image before adding to document
                        optimized_path = self.optimize_image_for_word(image_path)
                        
                        run = image_para.add_run()
                        run.add_picture(optimized_path, width=Inches(3.23), height=Inches(2.43))
                    except Exception as e:
                        image_para.text = f"[No image found for Tree {tree_number}]"
                        for run in image_para.runs:
                            run.italic = True
                else:
                    image_para.text = f"[No image found for Tree {tree_number}]"
                    for run in image_para.runs:
                        run.italic = True
                
                # Text cell
                text_cell = table.cell(row_index + 1, col_index)
                text_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                
                p = text_cell.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.space_after = p.space_before = Pt(0)
                
                # Combine runs to reduce operations
                text_content = f"TREE NUMBER : {tree_number}\nSPECIES      : {species}"
                tree_run = p.add_run(text_content)
                tree_run.font.name = arial_11_bold['name']
                tree_run.font.size = arial_11_bold['size']
                tree_run.font.bold = arial_11_bold['bold']
        
        # Clear memory
        if hasattr(self, 'match_cache'):
            self.match_cache.clear()
        
        # Save document
        file_name = os.path.join(
            self.output_folder_path, 
            f"TreeData_Batch{batch_num}_of_{total_batches}.docx"
        )
        doc.save(file_name)
        
        # Clean up temp files periodically
        if batch_num % 5 == 0 and hasattr(self, 'image_size_cache'):
            self.image_size_cache.clear()

    def update_ui(self, func):
        """Execute function in the main thread to update UI safely"""
        self.root.after(0, func)

    def run_in_background(self, func, callback=None):
        """Run a function in a background thread with optional callback"""
        def _worker():
            result = None
            error = None
            try:
                result = func()
            except Exception as e:
                error = e
            
            if callback:
                self.update_ui(lambda: callback(result, error))
        
        thread = threading.Thread(target=_worker)
        thread.daemon = True
        thread.start()
        return thread

if __name__ == "__main__":
    root = tk.Tk()
    app = TreeInventoryApp(root)
    root.mainloop()