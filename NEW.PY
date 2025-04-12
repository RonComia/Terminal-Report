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
        
        self.image_cache = {}

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

    def process_data(self):
        try:
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
            
            self.progress_frame.pack(fill=tk.X, pady=10)
            self.progress_bar["value"] = 0
            self.root.update()
            
            df = pd.read_excel(self.excel_file_path)
            
            if 'TREE NUMBER' not in df.columns and 'Tree Number' not in df.columns:
                if len(df.columns) >= 6:
                    tree_col_name = df.columns[5]
                    species_col_name = df.columns[6] if len(df.columns) > 6 else None
                    
                    df.rename(columns={tree_col_name: 'TREE NUMBER'}, inplace=True)
                    if species_col_name:
                        df.rename(columns={species_col_name: 'SPECIES'}, inplace=True)
                else:
                    messagebox.showerror(
                        "Column Error", 
                        "Excel file must contain 'TREE NUMBER' and 'SPECIES' columns or at least 6 columns"
                    )
                    self.progress_frame.pack_forget()
                    return
            elif 'Tree Number' in df.columns:
                df.rename(columns={'Tree Number': 'TREE NUMBER'}, inplace=True)
            
            if 'SPECIES' not in df.columns and 'Species' not in df.columns:
                if len(df.columns) >= 7:
                    species_col_name = df.columns[6]
                    df.rename(columns={species_col_name: 'SPECIES'}, inplace=True)
                else:
                    messagebox.showerror(
                        "Column Error", 
                        "Excel file must contain 'SPECIES' column or at least 7 columns"
                    )
                    self.progress_frame.pack_forget()
                    return
            elif 'Species' in df.columns:
                df.rename(columns={'Species': 'SPECIES'}, inplace=True)
            
            if self.start_row > 2:
                start_index = self.start_row - 1
                df = df.iloc[start_index:].reset_index(drop=True)
            
            self.progress_label.config(text="Loading images...")
            self.root.update()
            self.load_image_cache()
            
            total_rows = len(df)
            num_batches = (total_rows + self.batch_size - 1) // self.batch_size
            
            if self.num_batches_to_process is not None:
                num_batches = min(num_batches, self.num_batches_to_process)
            
            self.progress_bar["maximum"] = num_batches
            
            for batch_num in range(num_batches):
                self.progress_label.config(text=f"Processing batch {batch_num + 1} of {num_batches}...")
                self.progress_bar["value"] = batch_num
                self.root.update()
                
                batch_start = batch_num * self.batch_size
                batch_end = min((batch_num + 1) * self.batch_size, total_rows)
                batch_df = df.iloc[batch_start:batch_end]
                
                self.generate_word_doc(batch_df, batch_num + 1, num_batches)
                
                self.progress_bar["value"] = batch_num + 1
                self.root.update()
            
            self.progress_frame.pack_forget()
            messagebox.showinfo(
                "Processing Complete", 
                f"Generated {num_batches} Word documents in {self.output_folder_path}"
            )
            
        except Exception as e:
            self.progress_frame.pack_forget()
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def load_image_cache(self):
        """Load image cache with improved pattern matching to avoid incorrect substring matches"""
        self.image_cache = {}
        
        # List all image files in the directory and subdirectories
        image_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tif', '.tiff']
        all_files = []
        
        # Walk through all subdirectories
        for root, dirs, files in os.walk(self.image_folder_path):
            for file in files:
                if any(file.lower().endswith(ext) for ext in image_extensions):
                    all_files.append(os.path.join(root, file))
        
        # Show progress for loading images
        total_files = len(all_files)
        if total_files > 0:
            # Set up a progress dialog
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
                text=f"Loading image files (0/{total_files})...",
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
            
            # Process images
            for i, file_path in enumerate(all_files):
                if i % 10 == 0:
                    progress_label.config(text=f"Loading image files ({i+1}/{total_files})...")
                    image_progress["value"] = i + 1
                    progress_window.update()
                
                filename = os.path.basename(file_path)
                name_without_ext = os.path.splitext(filename)[0]
                
                # Store the original file for debugging
                self.image_cache[f"original_{name_without_ext}"] = file_path
                
                # Different matching strategies using word boundaries
                # Match exact numbers using word boundaries or start/end markers
                patterns = [
                    # Exact match for the full filename (no numbers around it)
                    r'^(\d+)$',
                    # Exact number at start with separator
                    r'^(\d+)[ _\-\.]', 
                    # Exact number with separator on both sides
                    r'[ _\-\.](\d+)[ _\-\.]',
                    # Exact number at end with separator
                    r'[ _\-\.](\d+)$'
                ]
                
                # Apply each pattern and store matches
                for pattern in patterns:
                    matches = re.findall(pattern, name_without_ext)
                    for match in matches:
                        try:
                            tree_num = int(match)
                            
                            # Store the exact number match (highest priority)
                            exact_key = f"exact_{tree_num}"
                            if exact_key not in self.image_cache:
                                self.image_cache[exact_key] = file_path
                                
                            # Store with the same format as in the file
                            formatted_key = f"formatted_{match}"
                            if formatted_key not in self.image_cache:
                                self.image_cache[formatted_key] = file_path
                        except ValueError:
                            pass
                
                # Also try to extract the primary number from the start
                primary_match = re.match(r'^(\d+)', name_without_ext)
                if primary_match:
                    num_str = primary_match.group(1)
                    try:
                        tree_num = int(num_str)
                        
                        # Use "start_" prefix to indicate this was at the start
                        self.image_cache[f"start_{tree_num}"] = file_path
                        
                        # Also store the original padded version
                        if len(num_str) > 1 and num_str.startswith('0'):
                            self.image_cache[f"padded_{num_str}"] = file_path
                    except ValueError:
                        pass
            
            progress_window.destroy()
            
            # Log the cache size for debugging
            print(f"Image cache loaded with {len(self.image_cache)} entries for {total_files} files")
            
            # Debug output
            print("Image Cache Keys (sample):")
            key_count = 0
            for key in list(self.image_cache.keys())[:20]:  # Show first 20 keys
                print(f"  {key} -> {os.path.basename(self.image_cache[key])}")
                key_count += 1
                if key_count >= 20:
                    break

    def get_image_path(self, tree_number):
        """Improved get_image_path that avoids substring matches"""
        if not tree_number or pd.isna(tree_number):
            return None
        
        try:
            # Clean the input
            tree_num_str = str(tree_number).strip()
            
            # Try to convert to integer
            try:
                tree_num = int(tree_num_str)
                # Create different formats for matching
                raw_num = str(tree_num)  # Without leading zeros
            except ValueError:
                # Not a valid integer, use as is
                raw_num = tree_num_str
            
            # Debug output
            print(f"Looking for image for tree number: {tree_num_str}")
            
            # ---- HIGHEST PRIORITY: EXACT MATCHES ----
            
            # 1. Try exact match key first (most precise)
            exact_key = f"exact_{tree_num}"
            if exact_key in self.image_cache:
                print(f"Found exact match: {os.path.basename(self.image_cache[exact_key])}")
                return self.image_cache[exact_key]
            
            # 2. Check for formatted exact match
            formatted_key = f"formatted_{tree_num_str}"
            if formatted_key in self.image_cache:
                print(f"Found formatted match: {os.path.basename(self.image_cache[formatted_key])}")
                return self.image_cache[formatted_key]
            
            # ---- SECONDARY PRIORITY: START-OF-FILENAME MATCHES ----
            
            # 3. Look for start-of-filename matches
            start_key = f"start_{tree_num}"
            if start_key in self.image_cache:
                # Verify this is a good match by checking the original filename
                path = self.image_cache[start_key]
                filename = os.path.basename(path)
                name_without_ext = os.path.splitext(filename)[0]
                
                # Extra validation to ensure it's not a false positive
                # Require exact boundary or separator after the number
                if re.match(f"^{tree_num}($|[ _\\-\\.])", name_without_ext):
                    print(f"Found start match: {filename}")
                    return path
                else:
                    print(f"Rejected false positive start match: {filename}")
            
            # 4. Try padded versions with the exact same padding length
            if len(tree_num_str) > 1 and tree_num_str.startswith('0'):
                padded_key = f"padded_{tree_num_str}"
                if padded_key in self.image_cache:
                    print(f"Found padded match: {os.path.basename(self.image_cache[padded_key])}")
                    return self.image_cache[padded_key]
            
            # ---- LAST RESORT: MANUAL SEARCH THROUGH CACHE ----
            
            # As a last resort, manually check filenames for exact matches
            potential_matches = []
            for key, path in self.image_cache.items():
                if key.startswith("original_"):
                    original_name = key[9:]  # Remove "original_" prefix
                    
                    # Look for the exact tree number with word boundaries
                    if re.search(f"(^|[ _\\-\\.])({tree_num})($|[ _\\-\\.])", original_name):
                        # Calculate match quality (lower is better)
                        # Prefer shorter filenames as they're likely more specific
                        match_quality = len(original_name)
                        potential_matches.append((match_quality, path))
            
            # Sort by match quality and return the best match
            if potential_matches:
                potential_matches.sort()
                best_match = potential_matches[0][1]
                print(f"Found manual match: {os.path.basename(best_match)}")
                return best_match
            
            print(f"No image found for tree number: {tree_num_str}")
            return None
            
        except Exception as e:
            print(f"Error matching tree {tree_number}: {str(e)}")
            return None

    def generate_word_doc(self, batch_df, batch_num, total_batches):
        doc = Document()
        
        sections = doc.sections
        for section in sections:
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.5)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)
        
        actual_rows = len(batch_df)
        trees_per_column = (actual_rows + 1) // 2
        
        table_rows = trees_per_column * 2
        
        table = doc.add_table(rows=table_rows, cols=2)
        
        table.style = 'Table Grid'
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        for cell in table.columns[0].cells:
            cell.width = Inches(3.25)
        for cell in table.columns[1].cells:
            cell.width = Inches(3.25)
        
        for i in range(table_rows):
            if i % 2 == 0:
                table.rows[i].height = Inches(2.5)
            else:
                table.rows[i].height = Inches(0.4)
        
        tree_index = 0
        
        for r in range(trees_per_column):
            for c in range(2):
                if tree_index < len(batch_df):
                    tree_row = batch_df.iloc[tree_index]
                    tree_number = str(tree_row['TREE NUMBER'])
                    species = str(tree_row['SPECIES']).upper()
                    
                    image_cell = table.cell(r * 2, c)
                    image_para = image_cell.paragraphs[0]
                    image_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    image_path = self.get_image_path(tree_number)
                    
                    if image_path and os.path.exists(image_path):
                        try:
                            run = image_para.add_run()
                            run.add_picture(image_path, width=Inches(3.23), height=Inches(2.43))
                        except Exception as e:
                            image_para.text = f"[No image found for Tree {tree_number}]"
                            for run in image_para.runs:
                                run.italic = True
                    else:
                        image_para.text = f"[No image found for Tree {tree_number}]"
                        for run in image_para.runs:
                            run.italic = True
                    
                    text_cell = table.cell(r * 2 + 1, c)
                    text_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    
                    text_cell.text = ""
                    p = text_cell.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    p.space_after = Pt(0)
                    p.space_before = Pt(0)
                    
                    tree_run = p.add_run(f"TREE NUMBER : {tree_number}")
                    tree_run.font.name = "Arial"
                    tree_run.font.size = Pt(11)
                    tree_run.font.bold = True
                    
                    p.add_run("\n")
                    
                    species_run = p.add_run(f"SPECIES      : {species}")
                    species_run.font.name = "Arial"
                    species_run.font.size = Pt(11)
                    species_run.font.bold = True
                    
                    tree_index += 1
                else:
                    table.cell(r * 2, c).text = ""
                    table.cell(r * 2 + 1, c).text = ""
        
        file_name = os.path.join(
            self.output_folder_path, 
            f"TreeData_Batch{batch_num}_of_{total_batches}.docx"
        )
        doc.save(file_name)

if __name__ == "__main__":
    root = tk.Tk()
    app = TreeInventoryApp(root)
    root.mainloop()