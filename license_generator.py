import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import json
import hashlib
import string
import random
from datetime import datetime, timedelta
from cryptography.fernet import Fernet
import wmi
import os

class LicenseGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("License Generator")
        self.root.geometry("600x500")
        self.root.configure(bg="#1E1E1E")
        
        self.create_widgets()
        
    def create_widgets(self):
        # Main frame
        main_frame = tk.Frame(self.root, bg="#1E1E1E", padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = tk.Label(
            main_frame,
            text="Tree Inventory - License Generator",
            bg="#1E1E1E",
            fg="#FFFFFF",
            font=("Arial", 14, "bold")
        )
        title_label.pack(pady=(0, 20))
        
        # Hardware ID display
        hw_frame = tk.Frame(main_frame, bg="#1E1E1E", pady=10)
        hw_frame.pack(fill=tk.X)
        
        hw_label = tk.Label(
            hw_frame,
            text="Hardware ID:",
            bg="#1E1E1E",
            fg="#FFFFFF",
            font=("Arial", 10)
        )
        hw_label.pack(anchor="w")
        
        self.hw_id = self.get_hardware_id()
        hw_id_label = tk.Label(
            hw_frame,
            text=self.hw_id,
            bg="#1E1E1E",
            fg="#FFC107",
            font=("Arial", 10, "bold")
        )
        hw_id_label.pack(anchor="w")
        
        # License options frame
        options_frame = tk.Frame(main_frame, bg="#1E1E1E", pady=10)
        options_frame.pack(fill=tk.X)
        
        # License duration
        duration_label = tk.Label(
            options_frame,
            text="License Duration (days):",
            bg="#1E1E1E",
            fg="#FFFFFF",
            font=("Arial", 10)
        )
        duration_label.pack(anchor="w")
        
        self.duration_var = tk.StringVar(value="365")
        duration_entry = tk.Entry(
            options_frame,
            textvariable=self.duration_var,
            bg="#2A2A2A",
            fg="#FFFFFF",
            insertbackground="#FFFFFF",
            width=10
        )
        duration_entry.pack(anchor="w")
        
        # License key display
        key_frame = tk.Frame(main_frame, bg="#1E1E1E", pady=10)
        key_frame.pack(fill=tk.BOTH, expand=True)
        
        key_label = tk.Label(
            key_frame,
            text="Generated License Key:",
            bg="#1E1E1E",
            fg="#FFFFFF",
            font=("Arial", 10)
        )
        key_label.pack(anchor="w")
        
        self.key_text = tk.Text(
            key_frame,
            height=5,
            bg="#2A2A2A",
            fg="#FFFFFF",
            insertbackground="#FFFFFF",
            font=("Arial", 12, "bold")
        )
        self.key_text.pack(fill=tk.X, pady=(5, 10))
        
        # Buttons frame
        button_frame = tk.Frame(main_frame, bg="#1E1E1E")
        button_frame.pack(fill=tk.X, pady=10)
        
        # Generate button
        generate_btn = tk.Button(
            button_frame,
            text="Generate License",
            bg="#1E1E1E",
            fg="#FFFFFF",
            activebackground="#2A2A2A",
            activeforeground="#FFFFFF",
            bd=1,
            relief=tk.SOLID,
            font=("Arial", 10, "bold"),
            command=self.generate_license
        )
        generate_btn.pack(side=tk.LEFT, padx=5)
        
        # Copy button
        copy_btn = tk.Button(
            button_frame,
            text="Copy to Clipboard",
            bg="#1E1E1E",
            fg="#FFFFFF",
            activebackground="#2A2A2A",
            activeforeground="#FFFFFF",
            bd=1,
            relief=tk.SOLID,
            font=("Arial", 10),
            command=self.copy_to_clipboard
        )
        copy_btn.pack(side=tk.LEFT, padx=5)
        
        # Save button
        save_btn = tk.Button(
            button_frame,
            text="Save License",
            bg="#1E1E1E",
            fg="#FFFFFF",
            activebackground="#2A2A2A",
            activeforeground="#FFFFFF",
            bd=1,
            relief=tk.SOLID,
            font=("Arial", 10),
            command=self.save_license
        )
        save_btn.pack(side=tk.LEFT, padx=5)
        
        # Status label
        self.status_label = tk.Label(
            main_frame,
            text="",
            bg="#1E1E1E",
            fg="#FFFFFF",
            font=("Arial", 10)
        )
        self.status_label.pack(pady=(10, 0))
        
    def get_hardware_id(self):
        """Generate a unique hardware ID based on system components"""
        try:
            c = wmi.WMI()
            system_info = c.Win32_ComputerSystemProduct()[0]
            cpu_info = c.Win32_Processor()[0]
            bios_info = c.Win32_BIOS()[0]
            
            hardware_str = f"{system_info.UUID}-{cpu_info.ProcessorId}-{bios_info.SerialNumber}"
            return hashlib.sha256(hardware_str.encode()).hexdigest()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate hardware ID: {str(e)}")
            return None
    
    def generate_license_key(self):
        """Generate a formatted license key"""
        chars = string.ascii_uppercase + string.digits
        key_parts = []
        for _ in range(5):
            part = ''.join(random.choices(chars, k=5))
            key_parts.append(part)
        return '-'.join(key_parts)
    
    def generate_license(self):
        """Generate a license file"""
        try:
            days = int(self.duration_var.get())
            if days < 1:
                raise ValueError("Duration must be positive")
            
            # Generate license key
            license_key = self.generate_license_key()
            
            # Create license data
            license_data = {
                'license_key': license_key,
                'hardware_id': self.hw_id,
                'expiry_date': (datetime.now() + timedelta(days=days)).isoformat(),
                'issue_date': datetime.now().isoformat()
            }
            
            # Create signature
            data_to_sign = json.dumps(license_data, sort_keys=True)
            signature = hashlib.sha256(data_to_sign.encode()).hexdigest()
            license_data['signature'] = signature
            
            # Display the license data
            self.key_text.delete("1.0", tk.END)
            self.key_text.insert("1.0", json.dumps(license_data, indent=2))
            
            self.status_label.config(text="License generated successfully", fg="#00FF00")
            
        except ValueError as e:
            self.status_label.config(text=f"Invalid input: {str(e)}", fg="#FF0000")
        except Exception as e:
            self.status_label.config(text=f"Error generating license: {str(e)}", fg="#FF0000")
    
    def copy_to_clipboard(self):
        """Copy the license data to clipboard"""
        try:
            license_text = self.key_text.get("1.0", tk.END).strip()
            self.root.clipboard_clear()
            self.root.clipboard_append(license_text)
            self.status_label.config(text="License copied to clipboard", fg="#00FF00")
        except Exception as e:
            self.status_label.config(text=f"Error copying to clipboard: {str(e)}", fg="#FF0000")
    
    def save_license(self):
        """Save the license to files"""
        try:
            # Generate encryption key
            key = Fernet.generate_key()
            fernet = Fernet(key)
            
            # Get license data
            license_text = self.key_text.get("1.0", tk.END).strip()
            license_data = json.loads(license_text)
            
            # Encrypt license data
            encrypted_data = fernet.encrypt(json.dumps(license_data).encode())
            
            # Ask for save location
            save_dir = filedialog.askdirectory(title="Select Directory to Save License Files")
            if not save_dir:
                return
            
            # Save files
            license_path = os.path.join(save_dir, 'license.key')
            key_path = os.path.join(save_dir, 'encryption.key')
            
            with open(license_path, 'wb') as f:
                f.write(encrypted_data)
            
            with open(key_path, 'wb') as f:
                f.write(key)
            
            self.status_label.config(
                text=f"License saved successfully to:\n{license_path}\n{key_path}",
                fg="#00FF00"
            )
            
        except Exception as e:
            self.status_label.config(text=f"Error saving license: {str(e)}", fg="#FF0000")

if __name__ == "__main__":
    root = tk.Tk()
    app = LicenseGenerator(root)
    root.mainloop() 