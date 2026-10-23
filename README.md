# Unbuncher: An Excel Cell Unmerger

Merged cells in Excel break the core functionality and usability of the spreadsheet. [Read more here](https://theexcelclub.medium.com/stop-do-not-merge-cells-in-excel-heres-why-with-fixes-the-excel-club-4318dccfd318).  
There is a much better alternative - centre across range; but if you receive an Excel file with merged cell, this tool will help.

Unbuncher is a simple tool to unmerge cells in Excel files and fill all previously merged cells with their content.

## How to Use

### Method 1: Drag and Drop (Easiest)
1. Simply drag your Excel file (.xlsx or .xlsm) onto `Excel_Unmerger.exe`
2. The tool will process your file automatically
3. A new file with `_nomerge` suffix will be created in the same folder

### Method 2: Command Line
1. Open Command Prompt or PowerShell
2. Navigate to the folder containing `Excel_Unmerger.exe`
3. Run: `Excel_Unmerger.exe "path\to\your\file.xlsx"`

## What It Does

- **Before**: Cells are merged and only the top-left cell contains data
- **After**: All cells are unmerged and each cell contains the original merged data

Example:
```
Before:  [  Merged Cell Content  ]  (spans 3 cells)
         [                      ]
         [                      ]

After:   [ Merged Cell Content ]
         [ Merged Cell Content ]  
         [ Merged Cell Content ]
```

## Supported File Formats

- Excel Workbook (.xlsx)
- Excel Macro-Enabled Workbook (.xlsm)

## Output

The tool creates a new file with the same name plus `_nomerge` suffix:
- Input: `MySpreadsheet.xlsx`
- Output: `MySpreadsheet_nomerge.xlsx`

## Requirements

- Windows (any recent version)
- No additional software needed - this is a standalone executable

## Troubleshooting

### "File not found" error
- Make sure the Excel file exists and the path is correct
- Avoid special characters in file names

### "Unsupported file format" error
- Only .xlsx and .xlsm files are supported
- Convert .xls files to .xlsx format first

### Permission errors
- Make sure you have write permissions in the folder
- Close the Excel file if it's open in Excel

## Need Help?

Contact your IT team or the person who provided this tool.

---
*Unbuncher: Excel Cell Unmerger v1.0*

----

## How to install

Simplest installation is to download the zip file - this contains the Windows exe file with everything needed to run. Just unzip and save on your machine; drag an excel file onto the exe to run. 
Unsigned .exe files may trigger virus warnings in windows: Excel_Unbuncher has been scanned by [VirusTotal](https://www.virustotal.com/gui/file/46c9829b67691f88bb6b72169a2d8f42901100043ba3ac92ecea0fa8ddb3b797?nocache=1)


### Build requirements (if not already installed):

To build; from command line install dependencies
   `pip install openpyxl pyinstaller`
Build the executable:
Double-click build_executable.bat or run it from command prompt
This will create a dist folder with Excel_Unmerger.exe
Optional: Add an icon file named excel.ico in the same folder for a nicer-looking executable

