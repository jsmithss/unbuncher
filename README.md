# Unbuncher: An Excel Cell Unmerger

Merged cells in Excel break the core functionality and usability of the spreadsheet. [Read more here](https://theexcelclub.medium.com/stop-do-not-merge-cells-in-excel-heres-why-with-fixes-the-excel-club-4318dccfd318)
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
