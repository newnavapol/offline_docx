# IPD Document Generator (Offline)
A standalone desktop application for generating hospital IPD documents.

## Build Instructions
Run the following command to build the executable (ensure you have PyInstaller installed):
```cmd
python -m PyInstaller --onefile --windowed --add-data "templates:templates" --name "IPD_Document_Generator" main.py
```
