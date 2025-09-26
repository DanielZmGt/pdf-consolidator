# app.py (version 4 - multi-source & open folder)

import tkinter
import customtkinter
from tkinter import filedialog, messagebox
import pandas as pd
import os
import shutil
import threading
import re
import sys
import subprocess

# Set the appearance of the app
customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("blue")

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # --------------- CONFIGURE WINDOW ---------------
        self.title("PDF Consolidator")
        self.geometry("700x550")

        # --------------- CONFIGURE GRID LAYOUT ---------------
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(4, weight=1)

        # --------------- VARIABLES TO STORE PATHS ---------------
        self.excel_path = tkinter.StringVar()
        self.source_folder_paths = [] # Changed to a list
        self.destination_folder_path = tkinter.StringVar()

        # --------------- CREATE WIDGETS ---------------

        # --- Excel File Selection ---
        self.excel_button = customtkinter.CTkButton(self, text="Select Excel File", command=self.select_excel_file)
        self.excel_button.grid(row=0, column=0, padx=20, pady=10)
        self.excel_label = customtkinter.CTkLabel(self, textvariable=self.excel_path, wraplength=450)
        self.excel_label.grid(row=0, column=1, padx=20, pady=10)

        # --- Source Folder Selection (Multiple) ---
        self.source_frame = customtkinter.CTkFrame(self) # Frame for buttons
        self.source_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        
        self.add_source_button = customtkinter.CTkButton(self.source_frame, text="Add Source Folder", command=self.add_source_folder)
        self.add_source_button.pack(side="top", fill="x", pady=(0, 5))
        
        self.clear_sources_button = customtkinter.CTkButton(self.source_frame, text="Clear Sources", command=self.clear_source_folders)
        self.clear_sources_button.pack(side="top", fill="x")
        
        self.source_textbox = customtkinter.CTkTextbox(self, height=80) # Textbox to list sources
        self.source_textbox.grid(row=1, column=1, padx=20, pady=10, sticky="nsew")
        self.source_textbox.configure(state="disabled")

        # --- Destination Folder Selection ---
        self.dest_button = customtkinter.CTkButton(self, text="Select Destination Folder", command=self.select_destination_folder)
        self.dest_button.grid(row=2, column=0, padx=20, pady=10)
        self.dest_label = customtkinter.CTkLabel(self, textvariable=self.destination_folder_path, wraplength=450)
        self.dest_label.grid(row=2, column=1, padx=20, pady=10)

        # --- Status Text Box ---
        self.status_textbox = customtkinter.CTkTextbox(self, height=120)
        self.status_textbox.grid(row=3, column=0, columnspan=2, padx=20, pady=10, sticky="nsew")
        self.status_textbox.insert("0.0", "Status updates will appear here...\n")
        self.status_textbox.configure(state="disabled")

        # --- Start Button ---
        self.start_button = customtkinter.CTkButton(self, text="Start Consolidating", command=self.start_consolidation_thread)
        self.start_button.grid(row=4, column=0, columnspan=2, padx=20, pady=20)


    def select_excel_file(self):
        path = filedialog.askopenfilename(title="Select the Excel file", filetypes=[("Excel Files", "*.xlsx *.xls")])
        self.excel_path.set(path)
        self.log(f"Excel file selected: {path}")

    def add_source_folder(self):
        path = filedialog.askdirectory(title="Select a folder containing PDFs")
        if path and path not in self.source_folder_paths:
            self.source_folder_paths.append(path)
            self.update_source_textbox()
            self.log(f"Added source folder: {path}")

    def clear_source_folders(self):
        self.source_folder_paths.clear()
        self.update_source_textbox()
        self.log("Cleared all source folders.")

    def update_source_textbox(self):
        self.source_textbox.configure(state="normal")
        self.source_textbox.delete("1.0", tkinter.END)
        if self.source_folder_paths:
            self.source_textbox.insert("1.0", "\n".join(self.source_folder_paths))
        self.source_textbox.configure(state="disabled")
    
    def select_destination_folder(self):
        path = filedialog.askdirectory(title="Select the folder to save copied PDFs")
        self.destination_folder_path.set(path)
        self.log(f"Destination folder selected: {path}")

    def log(self, message):
        self.status_textbox.configure(state="normal")
        self.status_textbox.insert(tkinter.END, message + "\n")
        self.status_textbox.see(tkinter.END)
        self.status_textbox.configure(state="disabled")

    def start_consolidation_thread(self):
        self.start_button.configure(state="disabled", text="Working...")
        process_thread = threading.Thread(target=self.consolidate_files)
        process_thread.start()

    def consolidate_files(self):
        try:
            excel_file = self.excel_path.get()
            source_folders = self.source_folder_paths
            dest_folder = self.destination_folder_path.get()

            if not excel_file or not source_folders or not dest_folder:
                self.log("ERROR: Please select an Excel file, at least one source folder, and a destination folder.")
                self.start_button.configure(state="normal", text="Start Consolidating")
                return

            self.log("\nStarting process...")
            df = pd.read_excel(excel_file, header=None)
            raw_values = df.iloc[:, 0].astype(str).str.strip().str.lower()
            
            self.log("Validating data from Excel file...")
            uuid_pattern = re.compile(r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$')
            invalid_format_mask = ~raw_values.str.match(uuid_pattern, na=False)
            duplicates_mask = raw_values.duplicated(keep=False)
            problem_mask = invalid_format_mask | duplicates_mask
            
            if problem_mask.any():
                problem_df = pd.DataFrame({'Problem Value': raw_values[problem_mask]})
                problem_df['Issue'] = ''
                problem_df.loc[invalid_format_mask[problem_mask], 'Issue'] = 'Invalid Format'
                problem_df.loc[duplicates_mask[problem_mask], 'Issue'] += ' Repeated Value'
                problem_df['Issue'] = problem_df['Issue'].str.strip()
                self.log(f"Found {len(problem_df)} problematic entries. Creating report sheet.")
                with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    problem_df.to_excel(writer, sheet_name='Validation Report', index=False)
            else:
                self.log("Data validation passed. No issues found.")

            valid_uuids = raw_values[~problem_mask]
            uuids_to_find = set(valid_uuids)
            self.log(f"Found {len(uuids_to_find)} valid and unique UUIDs to process.")
            os.makedirs(dest_folder, exist_ok=True)

            files_copied_count = 0
            uuids_found_in_source = set()

            for source_folder in source_folders:
                self.log(f"Scanning folder: {source_folder}")
                if not os.path.isdir(source_folder):
                    self.log(f"Warning: Source folder not found, skipping.")
                    continue
                for filename in os.listdir(source_folder):
                    if filename.lower().endswith('.pdf'):
                        file_uuid = os.path.splitext(filename)[0].lower()
                        if file_uuid in uuids_to_find:
                            source_path = os.path.join(source_folder, filename)
                            destination_path = os.path.join(dest_folder, filename)
                            if not os.path.exists(destination_path):
                                shutil.copy2(source_path, destination_path)
                                self.log(f"Copied: {filename}")
                                files_copied_count += 1
                            uuids_found_in_source.add(file_uuid)
            
            self.log(f"\n--- Process Complete ---")
            self.log(f"Total files copied: {files_copied_count}")
            
            unfound_uuids = uuids_to_find - uuids_found_in_source
            if unfound_uuids:
                self.log(f"Could not find PDFs for {len(unfound_uuids)} UUIDs.")

            if files_copied_count > 0:
                self.after(100, self.ask_to_open_folder, dest_folder)

        except Exception as e:
            self.log(f"An error occurred: {e}")
        finally:
            self.start_button.configure(state="normal", text="Start Consolidating")

    def ask_to_open_folder(self, folder_path):
        if messagebox.askyesno("Success!", "Process complete. Would you like to open the destination folder?"):
            try:
                if sys.platform == "win32":
                    os.startfile(os.path.realpath(folder_path))
                elif sys.platform == "darwin": # macOS
                    subprocess.run(["open", os.path.realpath(folder_path)])
                else: # linux
                    subprocess.run(["xdg-open", os.path.realpath(folder_path)])
            except Exception as e:
                self.log(f"Error opening folder: {e}")

if __name__ == "__main__":
    app = App()
    app.mainloop()