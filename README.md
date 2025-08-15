# Cell Analysis Database Auto-Update System

This system automates the process of copying the Cell_Log.xlsx file, filtering data by project, and appending new records to the master Cell Analysis Database.

## 🚀 Quick Start

### **First-Time Setup**

1. **Clone this repository**:
   ```bash
   git clone <repository-url>
   cd cell-analysis-db-update
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Create your configuration**:
   ```bash
   cp config.template.json config.json
   ```
   
4. **Edit `config.json`** with your actual file paths:
   ```json
   {
     "projects": ["YourProject1", "YourProject2"],
     "source_path": "/path/to/your/source_data/Cell_Log.xlsx",
     "work_dir": "/path/to/this/project/directory",
     "db_path": "/path/to/your/database.xlsx",
     "sheet_to_copy": "log",
     "target_sheet": "Slurry",
     "unique_key_col": "A",
     "logging_format": "[{timestamp}] {message}",
     "dry_run": true,
     "auto_backup": true
   }
   ```

5. **Test with sample data** (optional):
   - Use the provided `source_data/sample_Cell_Log.xlsx` and `source_data/sample_database.xlsx` for initial testing
   - Update your `config.json` to point to these sample files

## 📋 **Proper Source Data Management**

### **⚠️ Important: Working with Real Data**

When you have real cell analysis data to process:

1. **Add your data rows** to the Cell_Log.xlsx "log" sheet  
2. **Ensure Project column** contains your configured project names (case-sensitive)
3. **Save the file** - the system will auto-fetch from SharePoint Downloads
4. **Run the updater**: `python main.py`

### **🚨 Data Guidelines**

**DO:**
- ✅ Only add **real experimental data** that belongs in the database
- ✅ Use **actual project names** from your `config.json`
- ✅ Test with **dry-run mode first**: `python main.py --dry-run`

**DON'T:**
- ❌ Add fake/test data like "CELL001", "NEWCELL001", etc.
- ❌ Use made-up cell IDs for testing system functionality  
- ❌ Put test data in the production database

### **📊 Example of Real Data Entry**

```
UniqueID: ABC123
BatchNumber: BATCH2025_01  
Project: YourProject1
TestDate: 2025-01-14
Capacity_mAh: 2850
Voltage_V: 3.72
Temperature_C: 25
Notes: First production run
```

### **🔄 Typical Workflow**

1. **Conduct cell analysis experiments**
2. **Record results** in Cell_Log.xlsx (or download from SharePoint)
3. **Run dry-run mode** to preview changes: `python main.py --dry-run`  
4. **Check output folder** for results preview
5. **Switch to live mode** to update database: `python main.py --live`

### **Step 1: Get Latest Source Data**

**Option A: Manual Process**
1. **Download/copy** your Cell_Log.xlsx to the `source_data/` directory
2. **Ensure** it has a "log" sheet with your experimental data

**Option B: SharePoint Integration** (if applicable)
1. **Download** Cell_Log.xlsx from your SharePoint to Downloads folder
2. **Run copy script**: `python copy_sharepoint_file.py`
3. **Verify** source_data/Cell_Log.xlsx has current data

### **Step 2: Run Database Update**

1. **Run in dry-run mode first (recommended):**
   ```
   python main.py --dry-run
   ```

2. **Switch to live mode:**
   ```
   python main.py --live
   ```

3. **Get fresh data and run:**
   ```
   python main.py --get-sharepoint --dry-run
   ```

## 📋 Prerequisites

- Python 3.8+
- Required libraries: `pandas`, `openpyxl` (install via: `pip install -r requirements.txt`)
- Excel files in the expected locations
- Write permissions to work directory and database location

## 🎯 Version 2.0 - Refactored Architecture ✅ PRODUCTION READY

The system has been **completely refactored** from a single 585-line monolithic script into a clean, modular architecture with comprehensive safety features:

### **📦 Modular Components**
- **`main.py`** (258 lines) - Enhanced entry point with CLI options
- **`config.py`** (62 lines) - Configuration management and validation
- **`logging_utils.py`** (63 lines) - Logging setup and utilities  
- **`file_operations.py`** (112 lines) - File copying, versioning, backup
- **`data_processing.py`** (120 lines) - Data filtering and duplicate checking
- **`database.py`** (203 lines) - Database update operations and dry-run pipeline
- **`copy_sharepoint_file.py`** (133 lines) - SharePoint integration (enhanced)

### **✨ Enhanced Features**
- **Integrated SharePoint**: Single command to fetch and process data
- **Command-line options**: No config editing needed for mode switching
- **Safe by default**: Dry-run mode default, backups to output folder
- **Better error handling**: More informative debugging and validation
- **Comprehensive testing**: Dedicated test suite for modular architecture
- **Data contamination prevention**: Enhanced rules and test isolation

## ⚙️ Configuration

Edit `config.json` to customize the system:

```json
{
  "projects": ["Project-A", "Project-B", "Project-C", "Project-D", "Project-E"],
  "source_path": "/path/to/your/project/source_data/Cell_Log.xlsx",
  "work_dir": "/path/to/your/project",
  "db_path": "/path/to/your/database/Cell_Analysis_db.xlsx",
  "sheet_to_copy": "log",
  "target_sheet": "Slurry",
  "unique_key_col": "A",
  "logging_format": "[{timestamp}] {message}",
  "dry_run": true,
  "auto_backup": true
}
```

### Configuration Parameters

| Parameter | Description | Example |
|-----------|-------------|---------|
| `projects` | List of project names to filter by | `["Project-A", "Project-B"]` |
| `source_path` | Path to source Cell_Log.xlsx | `"source_data/Cell_Log.xlsx"` |
| `work_dir` | Working directory for copied files | Current project directory |
| `db_path` | Path to master database | Path to cellpy database |
| `sheet_to_copy` | Source sheet name | `"log"` |
| `target_sheet` | Target sheet name | `"Slurry"` |
| `unique_key_col` | Unique key column identifier | `"A"` |
| `dry_run` | Run in simulation mode | `true` or `false` |
| `auto_backup` | Create automatic backups | `true` or `false` |

## 🔄 How It Works

### **SharePoint Copy** (`copy_sharepoint_file.py`)
1. **Finds** newest Cell_Log.xlsx in Downloads folder
2. **Extracts** only the "log" sheet for optimal performance
3. **Saves** to source_data/Cell_Log.xlsx
4. **Validates** file size and data integrity

### **Database Update** (`main.py`)
1. **File Copy**: Creates a versioned copy of Cell_Log.xlsx (format: `YYMMDD_cellog_N.xlsx`)
2. **Filter Data**: Extracts rows matching specified projects from the "log" sheet
3. **Check Duplicates**: Compares unique keys against existing database entries
4. **Backup Database**: Creates timestamped backup (if enabled)
5. **Update Database**: Appends new rows to the "Slurry" sheet

### **Enhanced CLI Options**
- `python main.py` - Use config.json settings
- `python main.py --dry-run` - Force dry-run mode
- `python main.py --live` - Force live mode
- `python main.py --get-sharepoint` - Get fresh SharePoint data first

## 📊 Data Schema

### **Source File: Cell_Log.xlsx**
- **Sheet Name**: "log"
- **Structure**: Simple tabular format with 8 columns (A-H)
- **Key Columns**:
  - **Column A**: UniqueID (Primary Key) - Used for duplicate detection
  - **Column C**: Project - Used for filtering by project names
  - **Columns D-H**: TestDate, Capacity_mAh, Voltage_V, Temperature_C, Notes

### **Target Database: Cell_Analysis_db.xlsx**
- **Sheet Name**: "Slurry" 
- **Structure**: Complex multi-header format (cellpy system)
- **Header Structure**: 
  - Rows 1-4: Multi-level headers (field names, data types, etc.)
  - Row 5+: Actual data
- **Key Mappings**:
  - **Column A**: Unique key field (for duplicate detection)
  - **Column C**: Project field (matches source project column)

### **Business Rules**
- Duplicate keys (Column A) are automatically skipped
- Only rows matching projects in `config.json` are processed
- Existing data in database is preserved
- New data is appended at the end

## 📊 Data Flow

```
Source Cell_Log.xlsx
    ↓ (copy with versioning)
source_data/YYMMDD_cellog_N.xlsx
    ↓ (filter by projects)
Filtered Data
    ↓ (check for duplicates)
New Rows Only
    ↓ (backup database)
Database Backup Created
    ↓ (append to Slurry sheet)
Updated Database

# Dry-run mode saves to output/ folder instead
```

## 📝 Logging

All operations are logged to both console and `update_log.txt`:

```
[2025-01-13 09:14:49] Cell Analysis Database Update - Starting
[2025-01-13 09:14:49] Source file copied: 250813_cellog_1.xlsx
[2025-01-13 09:14:49] Read 127 rows from log sheet
[2025-01-13 09:14:49] Filtered to 5 rows:
[2025-01-13 09:14:49]   Project-A: 3 rows
[2025-01-13 09:14:49]   Project-B: 2 rows
[2025-01-13 09:14:49] Found 124 existing keys in database
[2025-01-13 09:14:49] New rows to add: 5
[2025-01-13 09:14:49] Database backup created: backup_database_20250113_091449.bak.xlsx
[2025-01-13 09:14:49] Successfully appended 5 rows to Slurry
```

## 🧪 Enhanced Dry-Run Mode

The system features an **enhanced dry-run mode** that executes the complete pipeline but saves all results to a safe `output/` folder:

### How Enhanced Dry-Run Works:
1. **Full Pipeline Execution**: Runs all steps including file copying, filtering, duplicate checking
2. **Safe Output Location**: All results saved to `output/` folder instead of production database
3. **Real File Operations**: Creates actual modified database file for inspection
4. **Complete Validation**: Test the entire workflow with real data transformations

### Output Files Created:
- `dryrun_[database]_[timestamp].xlsx` - Modified database with new data appended
- `backup_[database]_[timestamp].bak.xlsx` - Backup copy for reference
- `[date]_cellog_N.xlsx` - Versioned source file copy (in main directory)

### Example Dry-Run Output:
```
[2025-08-13 09:38:45] DRY RUN PIPELINE COMPLETED SUCCESSFULLY
[2025-08-13 09:38:45] Source file copied: 250813_cellog_4.xlsx
[2025-08-13 09:38:45] Rows filtered by projects: 3
[2025-08-13 09:38:45] New rows appended: 3 (Project-A:2, Project-B:1)
[2025-08-13 09:38:45] Output database: output/dryrun_database_20250813_093845.xlsx
[2025-08-13 09:38:45] ✅ All files saved to output/ folder - production database unchanged
```

### Benefits:
- **Complete Testing**: Validate entire workflow with real data
- **Safe Inspection**: Examine actual results before applying to production
- **Confidence Building**: See exactly what changes will be made
- **Debugging**: Troubleshoot data issues in safe environment

## 🧪 Testing

Run the test suite to validate functionality:

```bash
# Test all modules
python tests/test_refactored.py

# Or run from project root using unittest
python -m unittest tests.test_refactored
```

## 🚨 Safety Features

- **Enhanced Dry Run Mode**: Executes full pipeline and saves results to `output/` folder
- **Automatic Backups**: Creates timestamped backups before modifications
- **Duplicate Detection**: Prevents duplicate entries in database
- **Comprehensive Logging**: Detailed operation logs for troubleshooting
- **Error Handling**: Graceful handling of missing files, permissions, etc.

## ⚠️ Important Notes

- **Critical Database**: The target database is part of the cellpy system - never modify its structure
- **Always Test First**: Run in dry-run mode before live operations
- **Backup Strategy**: Ensure backups are working and recoverable
- **Permissions**: Verify write access to work directory and database location

## 📁 File Structure

```
cell-analysis-db-update/
├── main.py                     # Main entry point
├── setup.py                    # Quick setup script
├── config.template.json        # Configuration template
├── requirements.txt            # Dependencies  
├── README.md                   # This file
├── .gitignore                  # Git ignore rules
├── src/                        # Source code modules
│   ├── __init__.py             # Package initialization
│   ├── config.py               # Configuration management
│   ├── logging_utils.py        # Logging utilities
│   ├── file_operations.py      # File operations
│   ├── data_processing.py      # Data processing
│   ├── database.py             # Database operations
│   ├── copy_sharepoint_file.py # SharePoint file copy utility
│   └── cleanup_old_files.py    # Maintenance utility
├── tests/                      # Test suite
│   └── test_refactored.py      # Comprehensive test suite
├── source_data/                # Input data folder
│   ├── sample_Cell_Log.xlsx    # Sample source data for testing
│   ├── sample_database.xlsx    # Sample database for testing
│   ├── Cell_Log.xlsx           # Your actual source data (git-ignored)
│   └── [YYMMDD]_cellog_N.xlsx  # Versioned copies (created at runtime)
└── output/                     # Dry-run output folder (git-ignored)
    ├── dryrun_[database]_[timestamp].xlsx     # Modified database copy
    └── backup_[database]_[timestamp].bak.xlsx # Backup reference
```

## 🔧 Troubleshooting

### Common Issues

1. **"Source file is empty or corrupted"**
   - Run `python main.py --get-sharepoint --dry-run` to get fresh data
   - Check Downloads folder for latest Cell_Log.xlsx
   - Verify file was downloaded completely

2. **"Configuration file not found"**
   - Ensure `config.json` exists in the script directory
   - Copy from `config.template.json` and customize paths
   - Check file path and permissions

3. **"Source file not found"**
   - Run: `python main.py --get-sharepoint` to fetch fresh data
   - Check that source_data/Cell_Log.xlsx exists
   - Verify the file isn't open in Excel

4. **"Module not found" errors**
   - Ensure you're in the correct directory with all module files

5. **"Database file locked"**
   - Close Excel if the database is open
   - Check for other processes using the file

6. **"Permission denied"**
   - Run as administrator if needed
   - Check folder permissions for work directory

7. **"No rows to append"**
   - Verify project names match exactly (case-sensitive)
   - Check that source data contains the expected projects
   - Ensure source file has recent data (check file timestamp)

8. **"Permission denied" or "Process cannot access file"**
   - Close Excel if database file is open
   - Check if another process is using the file
   - Run as administrator if needed
   - Verify write permissions to work directory

9. **"Target sheet 'Slurry' not found"**
   - Open database file and verify sheet names
   - Check for typos in sheet name configuration
   - Ensure you're using the correct database file

### Diagnostic Commands

```python
# Test data access
import pandas as pd

# Check source file
df = pd.read_excel("source_data/Cell_Log.xlsx", sheet_name="log")
print(f"Source shape: {df.shape}")
print(f"Projects: {df.iloc[:, 2].unique()}")

# Check database
db_df = pd.read_excel("path/to/database.xlsx", sheet_name="Slurry", skiprows=4)
print(f"Database shape: {db_df.shape}")
```

### Emergency Recovery

**If Database is Corrupted:**
1. Stop all processes immediately
2. Restore from the most recent backup file (`*.bak.xlsx`)
3. Investigate root cause before resuming operations

**If Script Hangs:**
1. Kill the Python process
2. Check if database file is locked
3. Review log file for error messages
4. Restart with dry-run mode to diagnose

### Getting Help

- Check the log file: `update_log.txt`
- Run in dry-run mode: `python main.py --dry-run`
- Get fresh data: `python main.py --get-sharepoint --dry-run`
- Test the system: `python tests/test_refactored.py`
- Verify configuration settings

**Remember**: Always use dry-run mode first and maintain regular backups!

## 📈 Performance

- Typical execution time: 5-30 seconds (depending on data size)
- Memory usage: ~50-200 MB (depending on database size)
- Recommended for databases up to 10,000 rows

## 🔄 Automation

To run automatically, you can:

1. **Windows Task Scheduler**: Create a scheduled task
2. **Batch Script**: Create a `.bat` file with the Python command  
3. **PowerShell**: Use PowerShell scripts for more advanced automation

### Simple Batch Script Example:
```batch
@echo off
cd /d "C:\path\to\your\project"

REM Check if Python is available
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    pause
    exit /b 1
)

REM Check if required files exist
if not exist "main.py" (
    echo ERROR: main.py not found
    pause
    exit /b 1
)

REM Run the script
python main.py

REM Check exit code
if %errorlevel% equ 0 (
    echo Database update completed successfully!
    echo Check update_log.txt for detailed results
) else (
    echo ERROR: Database update failed!
    echo Check update_log.txt for error details
)

pause
```

### Safety Features for Automation:
- **Configuration validation** before execution
- **Safety confirmation** for live mode operations
- **Error handling** with informative messages
- **Automatic logging** for audit purposes

---

**Version**: 2.0 (Production Ready)  
**Author**: Open Source Community  
**Last Updated**: 2025-01-14  
**Status**: ✅ All systems operational, comprehensive testing complete  
**Documentation**: Consolidated into README.md for simplicity
