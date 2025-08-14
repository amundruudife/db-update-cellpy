# Cell Analysis Database Auto-Update System - Usage Guide

## 📋 **Proper Source Data Management**

### Cell_Log.xlsx - Your Source File

The `Cell_Log.xlsx` file in this directory should contain your **real cell analysis data**, not test data.

#### Current State:
- ✅ **File is properly structured** with correct headers
- ✅ **Contains no test data** (empty except headers)
- ✅ **Ready for your real data**

#### File Structure:
```
Sheet: "log"
Columns: UniqueID | BatchNumber | Project | TestDate | Capacity_mAh | Voltage_V | Temperature_C | Notes
```

### ⚠️ **Important: Adding Real Data**

When you have real cell analysis data to process:

1. **Open `Cell_Log.xlsx`**
2. **Add your data rows to the "log" sheet**
3. **Ensure Project column contains "Salamander", "SIS-Larger", or other configured projects**
4. **Save the file**
5. **Run the updater**: `python db_updater.py`

### 🚨 **Never Add Test Data**

- **DO NOT** add fake data like "CELL001", "NEWCELL001", etc.
- **DO NOT** use this for testing with made-up cell IDs
- **Only add real experimental data** that should be in the database

### 🧪 **For Testing System Functionality**

If you need to test the system:
1. **Keep dry_run: true** in config.json
2. **Add a single real data row** (if you have one)
3. **Check the output/ folder** for results
4. **Remove test row** after testing
5. **Switch to live mode** only for real data processing

### 📊 **Example of Real Data Entry**

```
UniqueID: ABC123
BatchNumber: BATCH2025_01
Project: Salamander
TestDate: 2025-08-13
Capacity_mAh: 2850
Voltage_V: 3.72
Temperature_C: 25
Notes: First production run
```

### 🔄 **Typical Workflow**

1. **Conduct cell analysis experiments**
2. **Record results in Cell_Log.xlsx**
3. **Run dry-run mode** to preview changes
4. **Switch to live mode** to update database
5. **Clear Cell_Log.xlsx** for next batch (optional)

---

## 🎯 **Remember**

**This system is designed to process real experimental data, not test data. The production database should only contain legitimate cell analysis results.**
