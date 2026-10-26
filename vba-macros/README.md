# VBA Macros for Warehouse Exit Requests ETL Pipeline

This folder contains VBA (Visual Basic for Applications) automation macros designed for processing and consolidating thousands of warehouse exit request files—including permanent, temporary, and consignment (custodial) records—directly within Microsoft Excel.

These macros serve as a robust alternative in restricted or offline environments where external Python packages cannot be installed.

## Core Automation Logic
Both macros follow a standardized 5-step ETL (Extract, Transform, Load) workflow to ensure consistency:

1. **Navigate:** Access the specified folder path containing the raw Excel files.
2. **Find:** Scan the directory to identify all valid Excel workbooks while filtering out temporary/lock files (e.g., `~$`).
3. **Open:** Sequentially open each file in a non-intrusive, read-only mode.
4. **Extract:** Read the relevant worksheet or data range from each file.
5. **Load:** Paste or stack the data into your target destination (either as a new sheet or appended to a master list).

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
  - **Sequential Stacking:** Data from File 2 is appended directly below File 1, and so on.
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
   
