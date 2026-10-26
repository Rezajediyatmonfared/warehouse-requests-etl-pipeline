# VBA Macros for Warehouse Exit Requests ETL Pipeline

This folder contains VBA (Visual Basic for Applications) automation macros designed for processing and consolidating thousands of warehouse exit request files—including permanent, temporary, and consignment (custodial) records—directly within Microsoft Excel.

These macros serve as a robust alternative in restricted or offline environments where external Python packages cannot be installed.

## Project Context & Purpose
In industrial warehouse management, tracking material movements requires handling thousands of individual exit requests. These records often arrive as fragmented Excel files, making manual aggregation impossible. These macros automate the consolidation of these records into structured, analysis-ready formats.

---

## Available Macros

### 1. `Module_MergeToSheets.bas`
- **Purpose:** Consolidates thousands of individual `.xlsx` warehouse exit files from a specified folder into a single workbook, placing each file into its own separate worksheet.
- **Key Features:**
  - Automatically skips hidden/veryHidden system sheets (such as `XLR_NoRangeSheet`).
  - Uses `Scripting.FileSystemObject` to safely handle Persian/Unicode folder paths and filenames.
  - Automatically sanitizes and generates unique sheet names.
  - Ignores temporary Excel lock files (`~$`).

### 2. `Module_MergeToMasterSheet.bas`
- **Purpose:** Merges all warehouse exit files sequentially into a single unified worksheet (`MasterData`), stacking each file's complete layout and headers one below the other.
- **Key Features:**
  - Preserves full headers (document numbers, dates, and warehouse tracking details) without data loss.
  - Inserts one blank row between records for clear visual separation.
  - Automatically configures right-to-left layout (`RTL`) for document standards.

---

## How to Use

1. Open Microsoft Excel and press **`ALT` + `F11`** to open the VBA Editor.
2. In the menu, go to **File > Import File...** (or right-click your VBAProject and select *Import File*).
3. Select either `.bas` file (`Module_MergeToSheets.bas` or `Module_MergeToMasterSheet.bas`).
4. Open the imported module and update the folder path constant at the top of the script:
```vba
   FolderPath = "D:\Your\Folder\Path\"
   
