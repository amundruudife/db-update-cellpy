#!/usr/bin/env python3
"""
Unit Test Suite for Refactored Cell Analysis Database Auto-Update System

This module contains comprehensive tests for all refactored modules.
"""

import unittest
import tempfile
import shutil
import json
import pandas as pd
from pathlib import Path
from datetime import datetime
import sys
import os

# Add the parent directory (project root) to the Python path to import src modules
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Import our refactored modules from src/
import src.config as config
import src.logging_utils as logging_utils
import src.file_operations as file_operations
import src.data_processing as data_processing
import src.database as database

class TestConfig(unittest.TestCase):
    """Test configuration module"""
    
    def setUp(self):
        """Set up test fixtures"""
        self.temp_dir = tempfile.mkdtemp()
        self.temp_path = Path(self.temp_dir)
        
    def tearDown(self):
        """Clean up test fixtures"""
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_valid_config_loading(self):
        """Test loading a valid configuration"""
        config_data = {
            "projects": ["TestProject"],
            "source_path": str(self.temp_path / "source.xlsx"),
            "work_dir": str(self.temp_path),
            "db_path": str(self.temp_path / "db.xlsx"),
            "sheet_to_copy": "log",
            "target_sheet": "Slurry",
            "unique_key_col": "A",
            "logging_format": "[{timestamp}] {message}",
            "dry_run": True,
            "auto_backup": True
        }
        
        # Create dummy files
        (self.temp_path / "source.xlsx").touch()
        (self.temp_path / "db.xlsx").touch()
        
        config_file = self.temp_path / "test_config.json"
        with open(config_file, 'w') as f:
            json.dump(config_data, f)
        
        # Test loading
        loaded_config = config.load_config(str(config_file))
        self.assertEqual(loaded_config['projects'], ["TestProject"])
        self.assertTrue(loaded_config['dry_run'])
    
    def test_missing_config_file(self):
        """Test error handling for missing config file"""
        with self.assertRaises(FileNotFoundError):
            config.load_config("nonexistent_config.json")

class TestLoggingUtils(unittest.TestCase):
    """Test logging utilities module"""
    
    def setUp(self):
        """Set up test fixtures"""
        self.temp_dir = tempfile.mkdtemp()
        self.temp_path = Path(self.temp_dir)
        
    def tearDown(self):
        """Clean up test fixtures"""
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_setup_logging(self):
        """Test logging setup"""
        logger = logging_utils.setup_logging(str(self.temp_path), "[{timestamp}] {message}")
        self.assertIsNotNone(logger)
        self.assertEqual(logger.name, 'db_updater')
        
        # Test that log file is created
        log_file = self.temp_path / "update_log.txt"
        logger.info("Test message")
        self.assertTrue(log_file.exists())
    
    def test_get_logger(self):
        """Test getting logger instance"""
        # Set up logger first
        logging_utils.setup_logging(str(self.temp_path), "[{timestamp}] {message}")
        
        # Get logger
        logger = logging_utils.get_logger()
        self.assertIsNotNone(logger)
        self.assertEqual(logger.name, 'db_updater')

class TestFileOperations(unittest.TestCase):
    """Test file operations module"""
    
    def setUp(self):
        """Set up test fixtures"""
        self.temp_dir = tempfile.mkdtemp()
        self.temp_path = Path(self.temp_dir)
        
        # Set up logging for file operations
        logging_utils.setup_logging(str(self.temp_path), "[{timestamp}] {message}")
        
        # Create a simple test Excel file
        test_data = pd.DataFrame({
            'UniqueID': ['TEST001', 'TEST002'],
            'Project': ['TestProj', 'TestProj'],
            'Data': [100, 200]
        })
        
        self.source_file = self.temp_path / "test_source.xlsx"
        test_data.to_excel(self.source_file, sheet_name='log', index=False)
        
    def tearDown(self):
        """Clean up test fixtures"""
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_copy_log_sheet_versioning(self):
        """Test file copying with version increment"""
        # First copy
        copied_path1 = file_operations.copy_log_sheet(str(self.source_file), str(self.temp_path))
        self.assertTrue(Path(copied_path1).exists())
        
        # Second copy should have incremented version
        copied_path2 = file_operations.copy_log_sheet(str(self.source_file), str(self.temp_path))
        self.assertTrue(Path(copied_path2).exists())
        self.assertNotEqual(copied_path1, copied_path2)
        
        # Check filename format
        filename1 = Path(copied_path1).name
        filename2 = Path(copied_path2).name
        
        today = datetime.now().strftime("%y%m%d")
        self.assertTrue(filename1.startswith(today))
        self.assertTrue(filename2.startswith(today))
        self.assertIn("_cellog_1.xlsx", filename1)
        self.assertIn("_cellog_2.xlsx", filename2)
    
    def test_backup_db(self):
        """Test database backup functionality"""
        # Create a dummy database file
        db_file = self.temp_path / "test_db.xlsx"
        db_file.write_text("test content")
        
        # Test backup creation
        backup_path = file_operations.backup_db(str(db_file), auto_backup=True, work_dir=str(self.temp_path))
        self.assertIsNotNone(backup_path)
        self.assertTrue(Path(backup_path).exists())
        self.assertIn(".bak", backup_path)
        
        # Test auto_backup disabled
        backup_path_none = file_operations.backup_db(str(db_file), auto_backup=False, work_dir=str(self.temp_path))
        self.assertIsNone(backup_path_none)

class TestDataProcessing(unittest.TestCase):
    """Test data processing module"""
    
    def setUp(self):
        """Set up test fixtures"""
        self.temp_dir = tempfile.mkdtemp()
        self.temp_path = Path(self.temp_dir)
        
        # Set up logging
        logging_utils.setup_logging(str(self.temp_path), "[{timestamp}] {message}")
        
        # Create test data in cellpy format with headers
        header_rows = [
            ['TEST INFO', 'op', 'proj', 'experiment'],  # Row 1: Main headers
            ['key', 'op', 'proj', 'experiment'],        # Row 2: Field names
            ['int', 'str', 'str', 'str'],               # Row 3: Data types
            ['key', 'op', 'proj', 'experiment']         # Row 4: Repeat field names
        ]
        
        # Actual data rows
        data_rows = [
            ['KEY001', 'MOS', 'Project-A', 'EXP001'],
            ['KEY002', 'MOS', 'Project-B', 'EXP002'],
            ['KEY003', 'MOS', 'Project-C', 'EXP003'],
            ['KEY004', 'MOS', 'Project-A', 'EXP004']
        ]
        
        # Combine headers and data
        all_rows = header_rows + data_rows
        
        # Create DataFrame with all rows
        columns = ['UniqueID', 'BatchNumber', 'Project', 'Data']
        self.test_data = pd.DataFrame(all_rows, columns=columns)
        
        self.test_file = self.temp_path / "test_data.xlsx"
        self.test_data.to_excel(self.test_file, sheet_name='log', index=False)
        
    def tearDown(self):
        """Clean up test fixtures"""
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_filter_by_projects(self):
        """Test project filtering functionality"""
        projects = ['Project-A', 'Project-B']
        
        filtered_df = data_processing.filter_by_projects(
            str(self.test_file), 
            'log', 
            projects
        )
        
        self.assertEqual(len(filtered_df), 3)  # Should get 2 Project-A + 1 Project-B
        # Check project column (index 2) contains expected projects
        project_col = filtered_df.iloc[:, 2]  # Column index 2 is the project column
        self.assertTrue(all(proj in projects for proj in project_col))

    def test_filter_empty_projects(self):
        """Test filtering with empty project list"""
        filtered_df = data_processing.filter_by_projects(
            str(self.test_file), 
            'log', 
            []
        )
        
        self.assertEqual(len(filtered_df), 0)

class TestDatabase(unittest.TestCase):
    """Test database operations module"""
    
    def setUp(self):
        """Set up test fixtures"""
        self.temp_dir = tempfile.mkdtemp()
        self.temp_path = Path(self.temp_dir)
        
        # Set up logging
        logging_utils.setup_logging(str(self.temp_path), "[{timestamp}] {message}")
        
        # Create test data
        self.test_data = pd.DataFrame({
            'UniqueID': ['NEW001', 'NEW002'],
            'Project': ['TestProj', 'TestProj'],
            'Data': [100, 200]
        })
        
    def tearDown(self):
        """Clean up test fixtures"""
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_update_slurry_dry_run(self):
        """Test dry-run mode for update_slurry"""
        # Create a dummy database file
        db_file = self.temp_path / "test_db.xlsx"
        
        # Create simple workbook for testing
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = 'Slurry'
        ws.append(['header1', 'header2', 'header3'])
        wb.save(db_file)
        wb.close()
        
        # Test dry run
        rows_added = database.update_slurry(
            self.test_data,
            str(db_file),
            'Slurry',
            dry_run=True
        )
        
        self.assertEqual(rows_added, len(self.test_data))
    
    def test_update_slurry_no_rows(self):
        """Test update_slurry with empty dataframe"""
        empty_df = pd.DataFrame()
        
        rows_added = database.update_slurry(
            empty_df,
            "dummy_path",
            "Slurry",
            dry_run=True
        )
        
        self.assertEqual(rows_added, 0)

def run_tests():
    """Run all tests"""
    # Create test suite
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    
    # Add test classes
    suite.addTests(loader.loadTestsFromTestCase(TestConfig))
    suite.addTests(loader.loadTestsFromTestCase(TestLoggingUtils))
    suite.addTests(loader.loadTestsFromTestCase(TestFileOperations))
    suite.addTests(loader.loadTestsFromTestCase(TestDataProcessing))
    suite.addTests(loader.loadTestsFromTestCase(TestDatabase))
    
    # Run tests
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    return result.wasSuccessful()

if __name__ == "__main__":
    success = run_tests()
    sys.exit(0 if success else 1) 