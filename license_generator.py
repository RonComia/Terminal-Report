import tkinter as tk
from tkinter import messagebox, ttk
import json
import hashlib
from datetime import datetime, timedelta
from cryptography.fernet import Fernet
import wmi
import os

class LicenseGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("License Generator")
        self.root.geometry("400x300")
        self.root.configure(bg="#1E1E1E")
        
        self.create_widgets()
        
    def create_widgets(self):
        # Hardware ID display
        hw_frame = tk.Frame(self.root, bg="#1E1E1E", pady=10)
        hw_frame.pack(fill=tk.X, padx=20)
        
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
        
        # License duration
        duration_frame = tk.Frame(self.root, bg="#1E1E1E", pady=10)
        duration_frame.pack(fill=tk.X, padx=20)
        
        duration_label = tk.Label(
            duration_frame,
            text="License Duration (days):",
            bg="#1E1E1E",
            fg="#FFFFFF",
            font=("Arial", 10)
        )
        duration_label.pack(anchor="w")
        
        self.duration_var = tk.StringVar(value="365")
        duration_entry = tk.Entry(
            duration_frame,
            textvariable=self.duration_var,
            bg="#2A2A2A",
            fg="#FFFFFF",
            insertbackground="#FFFFFF",
            width=10
        )
        duration_entry.pack(anchor="w")
        
        # Generate button
        generate_btn = tk.Button(
            self.root,
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
        generate_btn.pack(pady=20)
        
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
    
    def generate_license(self):
        """Generate a license file"""
        try:
            days = int(self.duration_var.get())
            if days < 1:
                raise ValueError("Duration must be positive")
            
            # Create license data
            license_data = {
                'hardware_id': self.hw_id,
                'expiry_date': (datetime.now() + timedelta(days=days)).isoformat(),
                'issue_date': datetime.now().isoformat()
            }
            
            # Create signature
            data_to_sign = json.dumps(license_data, sort_keys=True)
            signature = hashlib.sha256(data_to_sign.encode()).hexdigest()
            license_data['signature'] = signature
            
            # Generate encryption key
            key = Fernet.generate_key()
            fernet = Fernet(key)
            
            # Encrypt license data
            encrypted_data = fernet.encrypt(json.dumps(license_data).encode())
            
            # Save files
            with open('license.key', 'wb') as f:
                f.write(encrypted_data)
            
            with open('encryption.key', 'wb') as f:
                f.write(key)
            
            messagebox.showinfo(
                "Success",
                "License generated successfully!\n\n"
                "Files created:\n"
                "- license.key\n"
                "- encryption.key"
            )
            
        except ValueError as e:
            messagebox.showerror("Invalid Input", str(e))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate license: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = LicenseGenerator(root)
    root.mainloop() 