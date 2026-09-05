# IPD Document Generator (Offline) v1.0.0

A professional, standalone, and completely offline desktop application designed for Emergency Room physicians to quickly generate structured Inpatient (IPD) documents.

## Features
- **Offline Generation:** Compiles Doctor Orders, Inpatient H&P, and Informed Consent forms locally.
- **Manual Entry Interface:** A tabbed, dark-mode form with standard defaults for Physical Exams and Consents. 
- **Automated Routing:** Automatically saves files into a neat `YYYY/MM/DD/HN_Diagnosis` folder structure.
- **Cross-Platform:** Available as both a Windows `.exe` and a macOS `.app`.

## Installation & Download
You do not need to install Python to use this! You can download the latest compiled application directly from the **[GitHub Releases](../../releases)** page.
- **Windows Users:** Download `IPD_Document_Generator.exe`
- **Mac Users:** Download `IPD_Document_Generator_Mac.zip` (Extract it and move the `.app` to your Applications folder)

## Local Development & Building
If you wish to edit the source code and compile the application yourself:

1. Ensure Python 3.9+ is installed.
2. Install PyInstaller:
   ```bash
   pip install pyinstaller
   ```
3. Run the build command for your OS:
   - **Windows:**
     ```cmd
     python -m PyInstaller --onefile --windowed --add-data "templates;templates" --name "IPD_Document_Generator" main.py
     ```
   - **Mac:**
     ```bash
     python3 -m PyInstaller --onefile --windowed --add-data "templates:templates" --name "IPD_Document_Generator" main.py
     ```
