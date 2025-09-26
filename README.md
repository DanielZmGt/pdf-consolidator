# PDF Consolidator

A simple desktop application to find and copy PDF files based on a list of UUIDs in an Excel file.

## Features

* **Excel UUID List**: Reads a list of UUIDs from the first column of any `.xlsx` or `.xls` file.

* **Multi-Folder Search**: Select one or more source folders to search for your PDF files.

* **Data Validation**: Automatically checks the Excel list for duplicates and values with an invalid UUID format.

* **Validation Report**: If any issues are found, it creates a new sheet named "Validation Report" in your original Excel file detailing the problematic entries.

* **Case-Insensitive Matching**: Correctly matches `ABC-123.pdf` even if the Excel entry is `abc-123`.

* **Open Destination**: After successfully copying files, it asks if you want to open the destination folder for immediate access.

## How to Use

1. Clone this repository or download the source code.

2. Create a Python virtual environment and activate it.

3. Install the necessary dependencies from the `requirements.txt` file:

   ```
   pip install -r requirements.txt
   ```

4. Run the application:

   ```
   python app.py
   ```

5. Use the buttons in the app:

   * Select your Excel file.

   * Click "Add Source Folder" for each folder you want to search.

   * Select a destination folder for the copied PDFs.

6. Click "Start Consolidating" and review the status updates in the log.

## How to Create a Standalone Executable (.exe)

To share this app with others who don't have Python, you can package it into a single `.exe` file using PyInstaller.

1. Install PyInstaller:

   ```
   pip install pyinstaller
   ```

2. Run the build command from your terminal in the project directory:

   ```
   pyinstaller --onefile --windowed --hidden-import=customtkinter app.py
   ```

3. Find your finished application, `app.exe`, inside the newly created `dist` folder.